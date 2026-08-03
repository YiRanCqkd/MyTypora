const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const test = require('node:test');

const { createAuthorizationService } = require('../lib/authorization.cjs');
const {
  MAX_CONTENT_BYTES,
  createFilesService,
  validateSavePayload,
} = require('../lib/files.cjs');
const { validateSettingUpdate } = require('../lib/settings.cjs');

function hasCode(expectedCode) {
  return (error) => {
    assert.equal(error.code, expectedCode);
    assert.match(error.message, new RegExp(`^${expectedCode}:`));
    return true;
  };
}

test('非法设置 key 返回稳定错误码', () => {
  assert.throws(
    () => validateSettingUpdate('__proto__', true),
    hasCode('ERR_INVALID_SETTING_KEY'),
  );
});

test('超过 20 MiB 的 UTF-8 内容返回稳定错误码', () => {
  const oversized = 'a'.repeat(MAX_CONTENT_BYTES + 1);

  assert.throws(
    () => validateSavePayload({ content: oversized }),
    hasCode('ERR_DOCUMENT_TOO_LARGE'),
  );
});

test('未授权路径返回稳定错误码', async () => {
  const authorization = createAuthorizationService();
  const files = createFilesService({
    authorization,
    dialog: {},
    enforceAuthorization: true,
    fs,
    getEventWindow: () => null,
  });

  await assert.rejects(
    () => files.openFromPath(
      { sender: { id: 7 } },
      { filePath: 'D:\\docs\\private.md' },
      'trace-auth',
    ),
    hasCode('ERR_PATH_NOT_AUTHORIZED'),
  );
});
