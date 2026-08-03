function getContentType(filePath, path) {
  const types = {
    '.css': 'text/css; charset=utf-8',
    '.gif': 'image/gif',
    '.html': 'text/html; charset=utf-8',
    '.ico': 'image/x-icon',
    '.jpeg': 'image/jpeg',
    '.jpg': 'image/jpeg',
    '.js': 'text/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.map': 'application/json; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8',
    '.png': 'image/png',
    '.svg': 'image/svg+xml',
    '.ttf': 'font/ttf',
    '.woff': 'font/woff',
    '.woff2': 'font/woff2',
  };
  return types[path.extname(filePath).toLowerCase()] || 'application/octet-stream';
}

function patchIndexHtml(html) {
  const script = html.match(
    /<script\s+type="module"\s+crossorigin\s+src="([^"]+)"><\/script>/i,
  );
  if (!script?.[1]) return html;
  const bootstrap = `/__unireadmd_bootstrap__.js?entry=${encodeURIComponent(script[1])}`;
  return html
    .replace(/<link\s+rel="modulepreload"[^>]*>\s*/gi, '')
    .replace(script[0], `<script src="${bootstrap}"></script>`);
}

function createRendererServer(options) {
  const { fsSync, http, isDev, log, path, rootDirectory } = options;
  let server = null;
  let serverUrl = '';

  async function start() {
    if (isDev || server) return;
    const distRoot = path.resolve(rootDirectory, 'frontend', 'dist');
    server = http.createServer((request, response) => {
      const requestUrl = request.url || '/';
      const requestPath = decodeURIComponent(requestUrl.split('?')[0] || '/');
      if (requestPath === '/__unireadmd_bootstrap__.js') {
        const entry = new URL(`http://127.0.0.1${requestUrl}`).searchParams.get('entry') || '';
        const source = [
          '(async () => {',
          `  const entry = ${JSON.stringify(entry)};`,
          '  try { await import(entry); }',
          "  catch (error) { console.error('[boot-import-failed]', error); }",
          '})();',
        ].join('\n');
        response.writeHead(200, {
          'Content-Type': 'text/javascript; charset=utf-8',
          'Cache-Control': 'no-cache',
        });
        response.end(source);
        return;
      }
      const relative = (requestPath === '/' ? 'index.html' : requestPath).replace(/^\/+/, '');
      const candidate = path.resolve(distRoot, relative);
      const safeCandidate = candidate.startsWith(distRoot) ? candidate : '';
      const target = safeCandidate && fsSync.existsSync(safeCandidate)
        ? safeCandidate
        : path.join(distRoot, 'index.html');
      try {
        if (!fsSync.statSync(target).isFile()) throw new Error('not a file');
        if (target.toLowerCase().endsWith(`${path.sep}index.html`)) {
          response.writeHead(200, {
            'Content-Type': 'text/html; charset=utf-8',
            'Cache-Control': 'no-cache',
          });
          response.end(patchIndexHtml(fsSync.readFileSync(target, 'utf8')));
          return;
        }
        response.writeHead(200, {
          'Content-Type': getContentType(target, path),
          'Cache-Control': 'no-cache',
        });
        fsSync.createReadStream(target).pipe(response);
      } catch {
        response.writeHead(404);
        response.end('Not Found');
      }
    });
    await new Promise((resolve, reject) => {
      server.once('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    const address = server.address();
    serverUrl = `http://127.0.0.1:${typeof address === 'object' ? address.port : 0}/`;
    log('info', 'renderer static server started', 'boot-server', { url: serverUrl });
  }

  function close() {
    if (server) server.close();
    server = null;
    serverUrl = '';
  }

  return {
    close,
    getUrl: () => serverUrl,
    start,
  };
}

module.exports = {
  createRendererServer,
  getContentType,
  patchIndexHtml,
};
