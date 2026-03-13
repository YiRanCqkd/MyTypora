import { readFile, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const perfDir = join(__dirname, '..', 'perf');
const baselinePath = join(perfDir, 'baseline.json');
const latestPath = join(perfDir, 'latest.json');
const reportPath = join(perfDir, 'report.md');

const THRESHOLDS = {
  coldStartMs: 25,
  open5mbMs: 25,
  scrollSimulationMs: 30,
  search5mbMs: 25,
};

const ABSOLUTE_LIMITS = {
  coldStartMs: 10,
  open5mbMs: 200,
  scrollSimulationMs: 5,
  search5mbMs: 20,
};

function formatDelta(latest, baseline) {
  if (baseline === 0) {
    return 0;
  }

  return Number((((latest - baseline) / baseline) * 100).toFixed(2));
}

async function readJson(filePath) {
  const text = await readFile(filePath, 'utf8');
  return JSON.parse(text);
}

function compareMetrics(baseline, latest) {
  const rows = [];
  let hasRegression = false;

  for (const metric of Object.keys(THRESHOLDS)) {
    const baseValue = baseline.metrics?.[metric];
    const latestValue = latest.metrics?.[metric];

    if (typeof baseValue !== 'number' || typeof latestValue !== 'number') {
      continue;
    }

    const deltaPct = formatDelta(latestValue, baseValue);
    const deltaAbs = Number((latestValue - baseValue).toFixed(2));
    const threshold = THRESHOLDS[metric];
    const absoluteLimit = ABSOLUTE_LIMITS[metric];
    const isRegression = deltaPct > threshold && deltaAbs > absoluteLimit;
    if (isRegression) {
      hasRegression = true;
    }

    rows.push({
      metric,
      baseline: baseValue,
      latest: latestValue,
      deltaPct,
      deltaAbs,
      threshold,
      absoluteLimit,
      status: isRegression ? 'REGRESSION' : 'OK',
    });
  }

  return {
    rows,
    hasRegression,
  };
}

function buildReportMarkdown(baseline, latest, rows) {
  const lines = [
    '# 前端性能对比报告',
    '',
    `- baseline 时间: ${baseline.generatedAt}`,
    `- latest 时间: ${latest.generatedAt}`,
    '',
    '| 指标 | baseline | latest | 变化(%) | 增量 | 阈值(%) | 绝对阈值 | 状态 |',
    '| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |',
  ];

  for (const row of rows) {
    lines.push(
      `| ${row.metric} | ${row.baseline} | ${row.latest} | ${row.deltaPct} | ` +
      `${row.deltaAbs} | ${row.threshold} | ${row.absoluteLimit} | ${row.status} |`,
    );
  }

  lines.push(
    '',
    '说明：当变化百分比大于阈值且指标为耗时型时，判定为回归。',
    '可通过设置环境变量 `PERF_APPROVE=1` 放行一次人工审批。',
  );

  return lines.join('\n');
}

async function main() {
  if (!existsSync(latestPath)) {
    throw new Error(`缺少 latest 性能数据: ${latestPath}`);
  }

  if (!existsSync(baselinePath)) {
    throw new Error(`缺少 baseline 性能数据: ${baselinePath}`);
  }

  const baseline = await readJson(baselinePath);
  const latest = await readJson(latestPath);
  const result = compareMetrics(baseline, latest);
  const report = buildReportMarkdown(baseline, latest, result.rows);

  await writeFile(reportPath, `${report}\n`, 'utf8');
  console.log(`[perf] report written: ${reportPath}`);

  if (result.hasRegression && process.env.PERF_APPROVE !== '1') {
    throw new Error('检测到性能回归，已阻断。若人工审批通过，请设置 PERF_APPROVE=1。');
  }
}

main().catch((error) => {
  console.error('[perf] regression check failed', error.message);
  process.exitCode = 1;
});
