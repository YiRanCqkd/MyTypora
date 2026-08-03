import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTraceId, logBridge } from './trace';

function installLogBridge(logEvent: ReturnType<typeof vi.fn>) {
  Object.defineProperty(window, 'nativeBridge', {
    configurable: true,
    value: {
      app: {
        logEvent,
      },
    },
  });
}

describe('trace utils', () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it(
    'Given randomUUID support, when creating a trace, then prefix and compact UUID are used',
    () => {
    vi.stubGlobal('crypto', {
      randomUUID: () => '12345678-90ab-cdef-1234-567890abcdef',
    });

    expect(createTraceId('open:file')).toBe('open-file-1234567890');
    },
  );

  it(
    'Given no randomUUID, when creating a trace, then deterministic fallback values are used',
    () => {
    vi.stubGlobal('crypto', {});
    vi.spyOn(Date, 'now').mockReturnValue(1234);
    vi.spyOn(Math, 'random').mockReturnValue(0.5);

    expect(createTraceId('')).toBe('ui-1234-50000');
    },
  );

  it(
    'Given a working native logger, when logging, then the structured payload is forwarded',
    async () => {
    const logEvent = vi.fn().mockResolvedValue({ ok: true });
    installLogBridge(logEvent);

    await logBridge('info', 'opened', 'trace-1', { format: 'md' });

    expect(logEvent).toHaveBeenCalledWith({
      detail: { format: 'md' },
      level: 'info',
      message: 'opened',
      traceId: 'trace-1',
    });
    },
  );

  it.each([
    ['error', 'error'],
    ['info', 'warn'],
  ] as const)(
    'Given logger failure at %s level, when logging, then console.%s records the failure',
    async (level, consoleMethod) => {
      const failure = new Error('log failed');
      installLogBridge(vi.fn().mockRejectedValue(failure));
      const output = vi.spyOn(console, consoleMethod).mockImplementation(() => {});

      await logBridge(level, 'opened', 'trace-2');

      expect(output).toHaveBeenCalledWith('[bridge-log-failed] opened', failure);
    },
  );
});
