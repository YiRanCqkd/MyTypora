import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { pathToFileURL } from 'node:url';

const rawColorPattern = /#(?:[0-9a-f]{3,8})\b|\b(?:rgb|hsl)a?\([^)]*\)/gi;

export function findRawColorViolations(source, file = '<memory>') {
  const withoutComments = source.replace(/\/\*[\s\S]*?\*\//g, '');
  const violations = [];

  for (const [index, line] of withoutComments.split(/\r?\n/).entries()) {
    const customProperty = /^\s*--[\w-]+\s*:/.test(line);
    if (customProperty) {
      continue;
    }

    for (const match of line.matchAll(rawColorPattern)) {
      violations.push({
        file,
        line: index + 1,
        value: match[0],
      });
    }
  }

  return violations;
}

export function checkDesignTokenFiles(files) {
  return files.flatMap((file) => {
    const resolved = path.resolve(file);
    return findRawColorViolations(fs.readFileSync(resolved, 'utf8'), resolved);
  });
}

if (process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url) {
  const files = process.argv.slice(2);
  if (files.length === 0) {
    console.error('ERR_DESIGN_TOKEN_ARGUMENT: 至少提供一个 CSS 文件');
    process.exitCode = 1;
  } else {
    const violations = checkDesignTokenFiles(files);
    for (const violation of violations) {
      console.error(
        `ERR_DESIGN_TOKEN ${violation.file}:${violation.line} raw-color=${violation.value}`,
      );
    }
    process.exitCode = violations.length === 0 ? 0 : 1;
  }
}
