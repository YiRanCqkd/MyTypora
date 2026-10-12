const path = require('node:path');
const { findDocumentImages } = require('./document-images.cjs');
const { isWithin } = require('./workspace-tools.cjs');

async function collectDocumentImages({ fs, assets, event, source, content, root, destination }) {
  const parent = await fs.realpath(destination);
  const allowedRoot = await fs.realpath(root);
  if (!isWithin(allowedRoot, parent)) throw new Error('请选择当前工作区内的图片整理目录');
  const images = findDocumentImages(content).filter((image) =>
    !/^(?:[a-z][a-z\d+.-]*:|\/\/|#)/i.test(image.reference));
  if (!images.length) return { content, count: 0 };
  const token = new URL(assets.issue(event, source, root)).pathname.split('/').at(-1);
  const resources = new Map();
  let bytes = 0;
  for (const image of images) {
    const reference = decodeURIComponent(image.reference.split(/[?#]/)[0]);
    if (resources.has(reference)) continue;
    const resource = await assets.readForOwner(event, token, reference);
    bytes += resource.content.length;
    if (bytes > 64 * 1024 * 1024) throw new Error('整理图片超过 64 MiB');
    resources.set(reference, resource);
  }
  // 使用唯一目录，失败只移除本次副本；原图可能被其他文档使用，始终保留。
  const directory = await fs.mkdtemp(path.join(parent, `${path.parse(source).name}-images-`));
  try {
    const replacements = new Map();
    let index = 0;
    for (const [reference, resource] of resources) {
      const extension = resource.contentType.split('/')[1]
        .replace('svg+xml', 'svg').replace('x-icon', 'ico');
      const target = path.join(directory, `image-${index++}.${extension}`);
      await fs.writeFile(target, resource.content, { flag: 'wx' });
      replacements.set(reference, path.relative(path.dirname(source), target).replace(/\\/g, '/'));
    }
    for (const image of [...images].sort((a, b) => b.from - a.from)) {
      const reference = replacements.get(decodeURIComponent(image.reference.split(/[?#]/)[0]));
      const uri = reference.split('/').map((part) => encodeURIComponent(part)).join('/')
        .replace(/\(/g, '%28').replace(/\)/g, '%29');
      const alt = image.alt.replace(/[[\]\\]/g, '\\$&');
      const title = image.title ? ` "${image.title.replace(/["\\]/g, '\\$&')}"` : '';
      content = content.slice(0, image.from) + `![${alt}](${uri}${title})` + content.slice(image.to);
    }
    return { content, count: resources.size };
  } catch (error) {
    const real = await fs.realpath(directory);
    if (path.dirname(real) !== parent) throw new Error('图片整理清理目录越界');
    await fs.rm(real, { recursive: true, force: true });
    throw error;
  }
}

module.exports = { collectDocumentImages };
