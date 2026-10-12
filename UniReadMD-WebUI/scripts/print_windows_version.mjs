import fs from 'node:fs';

const packageFile = new URL('../package.json', import.meta.url);
const { version } = JSON.parse(fs.readFileSync(packageFile, 'utf8'));
if (!/^\d+\.\d+\.\d+$/.test(version)) throw new Error('Windows 资源版本须为三段数字版本');
console.log(version + '.0');
