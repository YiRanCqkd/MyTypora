import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const frontendPackage = path.join(scriptDirectory, '..', 'frontend', 'package.json');
const frontendRequire = createRequire(pathToFileURL(frontendPackage));
const ts = frontendRequire('typescript');
const { parse } = frontendRequire('@vue/compiler-sfc');

export function countScannedLines(source) {
  const lineStarts = ts.computeLineStarts(source);
  const counted = new Set();
  const scanner = ts.createScanner(
    ts.ScriptTarget.Latest,
    true,
    ts.LanguageVariant.Standard,
    source,
  );

  for (let token = scanner.scan(); token !== ts.SyntaxKind.EndOfFileToken; token = scanner.scan()) {
    const start = scanner.getTokenPos();
    const end = scanner.getTextPos();
    const startLine = ts.computeLineAndCharacterOfPosition(lineStarts, start).line;
    const endLine = ts.computeLineAndCharacterOfPosition(
      lineStarts,
      Math.max(start, end - 1),
    ).line;

    for (let line = startLine; line <= endLine; line += 1) {
      const lineStart = lineStarts[line];
      const lineEnd = line + 1 < lineStarts.length ? lineStarts[line + 1] : source.length;
      const segment = source.slice(Math.max(start, lineStart), Math.min(end, lineEnd));
      if (segment.trim()) {
        counted.add(line);
      }
    }
  }

  return counted.size;
}

export function extractCountableSource(filePath, source) {
  if (path.extname(filePath).toLowerCase() !== '.vue') {
    return source;
  }

  const result = parse(source, { filename: filePath });
  if (result.errors.length > 0) {
    const detail = result.errors.map((error) => String(error)).join('; ');
    throw new Error(`ERR_SOURCE_PARSE ${filePath}: ${detail}`);
  }

  return [result.descriptor.script?.content, result.descriptor.scriptSetup?.content]
    .filter((content) => typeof content === 'string')
    .join('\n');
}

export function countFile(filePath) {
  const source = fs.readFileSync(filePath, 'utf8');
  return countScannedLines(extractCountableSource(filePath, source));
}

function parseArguments(arguments_) {
  let limit = 250;
  const files = [];

  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index];
    if (argument === '--limit') {
      const value = Number(arguments_[index + 1]);
      if (!Number.isInteger(value) || value <= 0) {
        throw new Error('ERR_SOURCE_SIZE_ARGUMENT: --limit 必须是正整数');
      }
      limit = value;
      index += 1;
    } else {
      files.push(argument);
    }
  }

  if (files.length === 0) {
    files.push(
      path.join(scriptDirectory, '..', 'frontend', 'src', 'App.vue'),
      path.join(scriptDirectory, '..', 'electron', 'main.cjs'),
    );
  }

  return { files, limit };
}

export function runSourceSizeCheck(arguments_) {
  const { files, limit } = parseArguments(arguments_);
  const results = files.map((file) => {
    const resolved = path.resolve(file);
    return {
      file: resolved,
      lines: countFile(resolved),
      limit,
    };
  });
  const failures = results.filter((result) => result.lines > result.limit);

  for (const result of results) {
    const prefix = result.lines > result.limit ? 'ERR_SOURCE_SIZE' : 'SOURCE_SIZE_OK';
    console.log(`${prefix} ${result.file} lines=${result.lines} limit=${result.limit}`);
  }

  return failures.length === 0 ? 0 : 1;
}

if (process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url) {
  try {
    process.exitCode = runSourceSizeCheck(process.argv.slice(2));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(message);
    process.exitCode = 1;
  }
}
