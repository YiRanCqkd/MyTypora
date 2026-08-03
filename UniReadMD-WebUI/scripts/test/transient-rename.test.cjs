const assert = require('node:assert/strict');
const test = require('node:test');
const {
  renameWithTransientRetry,
  retryDelaysMs,
} = require('../transient-rename.cjs');

const source = 'D:\\verify\\release\\electron\\win-unpacked.tmp';
const destination = 'D:\\verify\\release\\electron\\win-unpacked';

const createRenameError = (code) => Object.assign(
  new Error(`${code}: rename failed`),
  {
    code,
    syscall: 'rename',
  },
);

test('Given 瞬态目标 EPERM, when 重命名, then 在有限等待内成功', async () => {
  let calls = 0;
  const waits = [];

  await renameWithTransientRetry({
    source,
    destination,
    platform: 'win32',
    rename: async () => {
      calls += 1;
      if (calls < 3) {
        throw createRenameError('EPERM');
      }
    },
    wait: async (milliseconds) => {
      waits.push(milliseconds);
    },
  });

  assert.equal(calls, 3);
  assert.deepEqual(waits, retryDelaysMs.slice(0, 2));
});

test('Given 非目标错误, when 重命名, then 第一次立即透传', async () => {
  let calls = 0;
  const waits = [];
  const expected = createRenameError('ENOENT');

  await assert.rejects(
    renameWithTransientRetry({
      source,
      destination,
      platform: 'win32',
      rename: async () => {
        calls += 1;
        throw expected;
      },
      wait: async (milliseconds) => {
        waits.push(milliseconds);
      },
    }),
    expected,
  );

  assert.equal(calls, 1);
  assert.deepEqual(waits, []);
});

test('Given 持续目标 EPERM, when 重试预算耗尽, then 保留失败', async () => {
  let calls = 0;
  const waits = [];
  const expected = createRenameError('EPERM');

  await assert.rejects(
    renameWithTransientRetry({
      source,
      destination,
      platform: 'win32',
      rename: async () => {
        calls += 1;
        throw expected;
      },
      wait: async (milliseconds) => {
        waits.push(milliseconds);
      },
    }),
    expected,
  );

  assert.equal(calls, retryDelaysMs.length + 1);
  assert.deepEqual(waits, retryDelaysMs);
});

test('Given 非 Electron 目录 EPERM, when 重命名, then 不重试', async () => {
  let calls = 0;
  const expected = createRenameError('EPERM');

  await assert.rejects(
    renameWithTransientRetry({
      source: 'D:\\verify\\other.tmp',
      destination: 'D:\\verify\\other',
      platform: 'win32',
      rename: async () => {
        calls += 1;
        throw expected;
      },
      wait: async () => {
        assert.fail('非目标路径不得等待');
      },
    }),
    expected,
  );

  assert.equal(calls, 1);
});
