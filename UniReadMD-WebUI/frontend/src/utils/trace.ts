export function createTraceId(prefix = 'ui'): string {
  const normalizedPrefix = prefix.replace(/[^A-Za-z0-9_-]/g, '-').slice(0, 16) || 'ui';

  if (globalThis.crypto?.randomUUID) {
    const raw = globalThis.crypto.randomUUID().replace(/-/g, '').slice(0, 10);
    return `${normalizedPrefix}-${raw}`;
  }

  const fallback = `${Date.now()}-${Math.floor(Math.random() * 100000)}`;
  return `${normalizedPrefix}-${fallback}`;
}

export async function logBridge(
  level: 'info' | 'error',
  message: string,
  traceId: string,
  detail?: unknown,
): Promise<void> {
  try {
    await window.nativeBridge.app.logEvent({
      level,
      message,
      traceId,
      detail,
    });
  } catch (error) {
    const logText = `[bridge-log-failed] ${message}`;
    if (level === 'error') {
      console.error(logText, error);
    } else {
      console.warn(logText, error);
    }
  }
}
