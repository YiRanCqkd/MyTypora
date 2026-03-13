import fs from 'node:fs';
import path from 'node:path';

const distAssetsDir = path.resolve('dist', 'assets');
const maxKb = Number(process.env.FRONTEND_BUNDLE_BUDGET_KB || 1400);

function toKb(bytes) {
  return bytes / 1024;
}

function readAssets() {
  if (!fs.existsSync(distAssetsDir)) {
    throw new Error('dist/assets 目录不存在，请先执行 npm run build');
  }

  return fs
    .readdirSync(distAssetsDir)
    .filter((name) => name.endsWith('.js'))
    .map((name) => {
      const filePath = path.join(distAssetsDir, name);
      const sizeBytes = fs.statSync(filePath).size;

      return {
        name,
        sizeBytes,
      };
    })
    .sort((a, b) => b.sizeBytes - a.sizeBytes);
}

function main() {
  const assets = readAssets();
  if (assets.length === 0) {
    throw new Error('未找到 JS 产物');
  }

  const mainEntry = assets.find((asset) => asset.name.startsWith('index-')) || assets[0];
  const mainKb = toKb(mainEntry.sizeBytes);

  console.log(`[bundle-check] main chunk: ${mainEntry.name} (${mainKb.toFixed(2)} KB)`);
  console.log(`[bundle-check] budget: ${maxKb.toFixed(2)} KB`);

  if (mainKb > maxKb) {
    throw new Error(`主包体积超限：${mainKb.toFixed(2)} KB > ${maxKb.toFixed(2)} KB`);
  }

  console.log('[bundle-check] pass');
}

main();
