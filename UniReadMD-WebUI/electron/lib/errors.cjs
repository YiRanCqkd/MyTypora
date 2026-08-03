function createContractError(code, message) {
  const error = new Error(`${code}: ${message}`);
  error.code = code;
  return error;
}

function requireString(input, fieldName, maxLength = 4096) {
  if (typeof input !== 'string' || input.length === 0) {
    throw createContractError('ERR_INVALID_ARGUMENT', `${fieldName} 必须是非空字符串`);
  }
  if (input.length > maxLength) {
    throw createContractError('ERR_INVALID_ARGUMENT', `${fieldName} 超过长度限制`);
  }
  return input;
}

function safeErrorMessage(error) {
  return error instanceof Error ? `${error.name}: ${error.message}` : String(error);
}

module.exports = {
  createContractError,
  requireString,
  safeErrorMessage,
};
