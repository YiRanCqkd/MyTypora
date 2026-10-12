const path = require('node:path');
const MarkdownIt = require('markdown-it');

function findDocumentImages(content) {
  const markdown = new MarkdownIt({ html: true });
  const imageRule = markdown.inline.ruler.__rules__.find((rule) => rule.name === 'image').fn;
  markdown.inline.ruler.at('image', (state, silent) => {
    const from = state.pos;
    const matched = imageRule(state, silent);
    if (matched && !silent) {
      const image = state.tokens.at(-1);
      image.meta = { from, to: state.pos };
    }
    return matched;
  });
  const lines = content.split('\n');
  const offsets = [];
  let position = 0;
  for (const line of lines) { offsets.push(position); position += line.length + 1; }
  const records = [];
  for (const token of markdown.parse(content, {})) {
    if (token.type !== 'inline' || !token.map) continue;
    const mapping = [];
    let lineIndex = token.map[0];
    for (const line of token.content.split('\n')) {
      let column = -1;
      while (lineIndex < token.map[1]) {
        column = lines[lineIndex].indexOf(line);
        if (column >= 0) break;
        lineIndex += 1;
      }
      if (column < 0) break;
      for (let index = 0; index <= line.length; index += 1) {
        mapping.push(offsets[lineIndex] + column + index);
      }
      lineIndex += 1;
    }
    for (const image of token.children || []) {
      if (image.type !== 'image' || !image.meta) continue;
      const { from, to } = image.meta;
      if (mapping[from] === undefined || mapping[to - 1] === undefined) {
        throw new Error('图片源码映射失败，请先将复杂图片引用改为独立行');
      }
      records.push({ from: mapping[from], to: mapping[to - 1] + 1,
        reference: image.attrGet('src'), alt: image.content,
        title: image.attrGet('title') || '' });
    }
  }
  return records;
}

async function relocateDocumentImages({ content, source, target, event, assets, fs, rootDirectory }) {
  if (!assets || path.dirname(source) === path.dirname(target)) {
    return { content, cleanup: async () => {} };
  }
  const images = findDocumentImages(content).filter((image) => {
    return !/^(?:[a-z][a-z\d+.-]*:|\/\/|#)/i.test(image.reference);
  });
  if (!images.length) return { content, cleanup: async () => {} };
  const token = new URL(assets.issue(event, source, rootDirectory)).pathname.split('/').at(-1);
  const resources = new Map();
  let total = 0;
  for (const image of images) {
    const reference = decodeURIComponent(image.reference.split(/[?#]/)[0]);
    if (resources.has(reference)) continue;
    const resource = await assets.readForOwner(event, token, reference);
    total += resource.content.length;
    if (total > 64 * 1024 * 1024) throw new Error('迁移图片超过 64 MiB');
    resources.set(reference, resource);
  }
  // 只创建目标文档旁的唯一目录，不覆盖用户已有资源，也不删除原图。
  const parent = await fs.realpath(path.dirname(target));
  const directory = await fs.mkdtemp(path.join(parent, `${path.parse(target).name}-assets-`));
  const cleanup = async () => {
    const real = await fs.realpath(directory);
    if (path.dirname(real) !== parent) throw new Error('资源清理目录越界');
    await fs.rm(real, { recursive: true, force: true });
  };
  try {
    const replacements = new Map();
    let index = 0;
    for (const [reference, resource] of resources) {
      const extension = resource.contentType.split('/')[1].replace('svg+xml', 'svg');
      const filePath = path.join(directory, `image-${index++}.${extension}`);
      await fs.writeFile(filePath, resource.content, { flag: 'wx' });
      replacements.set(reference, path.relative(parent, filePath).replace(/\\/g, '/'));
    }
    for (const image of [...images].sort((a, b) => b.from - a.from)) {
      const reference = replacements.get(decodeURIComponent(image.reference.split(/[?#]/)[0]));
      const alt = image.alt.replace(/[[\]\\]/g, '\\$&');
      const title = image.title ? ` "${image.title.replace(/["\\]/g, '\\$&')}"` : '';
      const uri = encodeURI(reference).replace(/[()]/g, (char) => encodeURIComponent(char)
        === char ? `%${char.charCodeAt(0).toString(16)}` : encodeURIComponent(char));
      content = content.slice(0, image.from) + `![${alt}](${uri}${title})`
        + content.slice(image.to);
    }
    const frontMatter = /^---\r?\n([\s\S]*?)\r?\n---/.exec(content);
    if (frontMatter) {
      const metadata = frontMatter[1].split(/\r?\n/)
        .filter((line) => !/^typora-(?:root-url|copy-images-to):/.test(line));
      metadata.push(`typora-copy-images-to: ${JSON.stringify(path.basename(directory))}`);
      content = content.replace(frontMatter[0], `---\n${metadata.join('\n')}\n---`);
    }
    return { content, cleanup };
  } catch (error) { await cleanup(); throw error; }
}

module.exports = { findDocumentImages, relocateDocumentImages };
