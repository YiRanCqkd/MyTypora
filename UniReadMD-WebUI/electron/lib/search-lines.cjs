const { Worker, isMainThread, parentPort, workerData } = require('node:worker_threads');

function buildPattern(query, options = {}) {
  const escaped = options.regularExpression ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const boundary = '[\\p{L}\\p{N}_]';
  const source = options.wholeWord
    ? `(?<!${boundary})(?:${escaped})(?!${boundary})` : escaped;
  return new RegExp(source, options.caseSensitive ? 'u' : 'iu');
}

async function searchLines(lines, query, options = {}, limit = 201) {
  if (!options.regularExpression && !options.wholeWord) {
    const needle = options.caseSensitive ? query : query.toLowerCase();
    const indices = [];
    for (let index = 0; index < lines.length; index += 1) {
      const text = options.caseSensitive ? lines[index] : lines[index].toLowerCase();
      if (text.includes(needle)) indices.push(index);
      if (indices.length >= limit) break;
    }
    return indices;
  }
  const expression = buildPattern(query, options);
  // 用户正则在独立线程执行，灾难性回溯不得阻塞窗口和保存操作。
  return new Promise((resolve, reject) => {
    const workerPath = __filename.replace(/([\\/])app\.asar(?=[\\/])/,
      '$1app.asar.unpacked');
    const worker = new Worker(workerPath, { workerData: {
      lines, source: expression.source, flags: expression.flags, limit,
    } });
    const timer = setTimeout(() => {
      void worker.terminate();
      reject(new Error('正则搜索超时，请简化表达式后重试'));
    }, 1000);
    worker.once('message', (result) => { clearTimeout(timer); resolve(result); });
    worker.once('error', (error) => { clearTimeout(timer); reject(error); });
    worker.once('exit', (code) => {
      clearTimeout(timer);
      if (code !== 0) reject(new Error('搜索线程已终止'));
    });
  });
}

if (!isMainThread) {
  const expression = new RegExp(workerData.source, workerData.flags);
  const indices = [];
  for (let index = 0; index < workerData.lines.length; index += 1) {
    if (expression.test(workerData.lines[index])) indices.push(index);
    if (indices.length >= workerData.limit) break;
  }
  parentPort.postMessage(indices);
}

module.exports = { searchLines, buildPattern };
