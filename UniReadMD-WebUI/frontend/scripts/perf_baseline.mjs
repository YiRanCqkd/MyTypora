import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const perfDir = join(__dirname, '..', 'perf');
const latestPath = join(perfDir, 'latest.json');
const baselinePath = join(perfDir, 'baseline.json');

function extractHeadings(content) {
  const lines = content.split(/\r?\n/);
  const output = [];

  lines.forEach((line, index) => {
    const match = /^(#{1,6})\s+(.+)$/.exec(line.trim());
    if (!match) {
      return;
    }

    output.push({
      line: index + 1,
      level: match[1].length,
      text: match[2],
    });
  });

  return output;
}

function findActiveHeadingLine(headings, cursorLine) {
  let matchedLine = 0;

  for (const heading of headings) {
    if (heading.line <= cursorLine) {
      matchedLine = heading.line;
      continue;
    }

    break;
  }

  return matchedLine;
}

function findMatchedLines(content, query, maxResults = 200) {
  const normalized = query.trim().toLowerCase();
  if (!normalized) {
    return [];
  }

  return content
    .split(/\r?\n/)
    .map((line, index) => ({
      lineNo: index + 1,
      line,
    }))
    .filter((item) => item.line.toLowerCase().includes(normalized))
    .slice(0, maxResults);
}

function buildLargeMarkdown(targetBytes = 5 * 1024 * 1024) {
  const blocks = [];
  let estimatedBytes = 0;
  let index = 0;

  while (estimatedBytes < targetBytes) {
    index += 1;
    const block =
      `# Section ${index}\n\n` +
      `This is a benchmark paragraph for UniReadMD.\n\n` +
      '```mermaid\n' +
      `graph TD\nA${index}-->B${index}\n` +
      '```\n\n' +
      `Keyword-line-${index}-search-target\n\n` +
      '$$E = mc^2$$\n';

    blocks.push(block);
    estimatedBytes += Buffer.byteLength(block, 'utf8');
  }

  return blocks.join('\n');
}

function toFixed(value) {
  return Number(value.toFixed(2));
}

function getMedian(values) {
  if (values.length === 0) {
    return 0;
  }

  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 1) {
    return sorted[middle];
  }

  return (sorted[middle - 1] + sorted[middle]) / 2;
}

function measureMedian(times, runner) {
  const samples = [];
  for (let index = 0; index < times; index += 1) {
    const begin = performance.now();
    runner();
    samples.push(performance.now() - begin);
  }

  return toFixed(getMedian(samples));
}

async function runBenchmarks() {
  const mdModule = await import('markdown-it');
  const MarkdownIt = mdModule.default;
  const document5mb = buildLargeMarkdown();
  const coldStartDocument = '# boot\n\n' + 'bootstrap line\n\n'.repeat(3000);

  const coldStartMs = measureMedian(5, () => {
    const markdown = new MarkdownIt({
      html: false,
      linkify: true,
      breaks: true,
    });
    markdown.render(coldStartDocument);
  });

  const open5mbMs = measureMedian(3, () => {
    const markdown = new MarkdownIt({
      html: false,
      linkify: true,
      breaks: true,
    });
    markdown.render(document5mb);
  });

  const headings = extractHeadings(document5mb);
  const scrollSimulationMs = measureMedian(5, () => {
    for (let line = 1; line <= 20000; line += 11) {
      findActiveHeadingLine(headings, line);
    }
  });

  const search5mbMs = measureMedian(5, () => {
    findMatchedLines(document5mb, 'search-target', 1000);
  });

  return {
    generatedAt: new Date().toISOString(),
    documentSizeBytes: Buffer.byteLength(document5mb, 'utf8'),
    metrics: {
      coldStartMs,
      open5mbMs,
      scrollSimulationMs,
      search5mbMs,
    },
  };
}

async function main() {
  const result = await runBenchmarks();
  await mkdir(perfDir, {
    recursive: true,
  });
  await writeFile(latestPath, `${JSON.stringify(result, null, 2)}\n`, 'utf8');

  if (process.argv.includes('--update-baseline')) {
    await writeFile(baselinePath, `${JSON.stringify(result, null, 2)}\n`, 'utf8');
    console.log(`[perf] baseline updated: ${baselinePath}`);
  }

  console.log(`[perf] latest result written: ${latestPath}`);
  console.log(JSON.stringify(result.metrics, null, 2));
}

main().catch((error) => {
  console.error('[perf] benchmark failed', error);
  process.exitCode = 1;
});
