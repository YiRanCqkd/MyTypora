import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { pathToFileURL } from 'node:url';

export class E2eBaselineError extends Error {
  constructor(code, message) {
    super(`${code}: ${message}`);
    this.name = 'E2eBaselineError';
    this.code = code;
  }
}

function collectResults(suites, parents = [], results = []) {
  for (const suite of suites ?? []) {
    const nextParents = suite.title && suite.title !== suite.file
      ? [...parents, suite.title]
      : parents;
    for (const spec of suite.specs ?? []) {
      for (const test of spec.tests ?? []) {
        const status = test.results?.at(-1)?.status ?? 'unknown';
        results.push({
          status,
          id: `[${test.projectName}] › ${spec.file}:${spec.line}:${spec.column} › ${[
            ...nextParents,
            spec.title,
          ].filter(Boolean).join(' › ')}`,
        });
      }
    }
    collectResults(suite.suites, nextParents, results);
  }
  return results;
}

export function evaluateE2eBaseline(report, baseline) {
  if (!Array.isArray(baseline.knownFailureIds) || !Number.isInteger(baseline.minimumPassed)) {
    throw new E2eBaselineError('ERR_E2E_BASELINE_SCHEMA', '基线结构无效');
  }

  const results = collectResults(report.suites);
  const failures = results.filter((result) => result.status === 'failed');
  const passed = results.filter((result) => result.status === 'passed').length;
  const skipped = results.filter((result) => result.status === 'skipped').length;
  const known = new Set(baseline.knownFailureIds);
  const unexpected = failures.filter((failure) => !known.has(failure.id));

  if ((report.stats?.flaky ?? 0) > 0) {
    throw new E2eBaselineError('ERR_E2E_FLAKY', `flaky=${report.stats.flaky}`);
  }
  if (skipped > 0 || (report.stats?.skipped ?? 0) > 0) {
    throw new E2eBaselineError('ERR_E2E_SKIPPED', `skipped=${skipped}`);
  }
  if (unexpected.length > 0) {
    throw new E2eBaselineError(
      'ERR_E2E_NEW_FAILURE',
      unexpected.map((failure) => failure.id).join(' | '),
    );
  }
  if (passed < baseline.minimumPassed) {
    throw new E2eBaselineError(
      'ERR_E2E_PASS_FLOOR',
      `passed=${passed} minimum=${baseline.minimumPassed}`,
    );
  }

  return {
    passed,
    failed: failures.length,
    skipped,
    knownFailures: baseline.knownFailureIds.length,
  };
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(path.resolve(file), 'utf8'));
}

function parseArguments(arguments_) {
  const values = new Map();
  for (let index = 0; index < arguments_.length; index += 2) {
    values.set(arguments_[index], arguments_[index + 1]);
  }
  const report = values.get('--report') || process.env.PLAYWRIGHT_JSON_OUTPUT_NAME;
  const baseline = values.get('--baseline');
  if (!report || !baseline) {
    throw new E2eBaselineError(
      'ERR_E2E_ARGUMENT',
      '必须提供 --report 和 --baseline',
    );
  }
  return { baseline, report };
}

if (process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url) {
  try {
    const files = parseArguments(process.argv.slice(2));
    const result = evaluateE2eBaseline(readJson(files.report), readJson(files.baseline));
    console.log(`E2E_BASELINE_OK ${JSON.stringify(result)}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(message);
    process.exitCode = 1;
  }
}
