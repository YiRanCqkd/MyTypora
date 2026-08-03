import { createTraceId, logBridge } from '../../utils/trace';
import type { useApplicationStatus } from '../editor/application-status';
type UseTaskRunnerDependencies =
  ReturnType<typeof useApplicationStatus>;

export function useTaskRunner(input: UseTaskRunnerDependencies) {
  const { errorMessage,
  } = input;

  async function runTask<T>(
    taskName: string,
    executor: (traceId: string) => Promise<T>,
    options?: {
      fallbackHint?: string;
    },
  ): Promise<T | null> {
    const traceId = createTraceId(taskName);

    try {
      await logBridge('info', `${taskName}:start`, traceId);
      const result = await executor(traceId);
      await logBridge('info', `${taskName}:ok`, traceId);
      errorMessage.value = '';

      return result;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const fallbackText = options?.fallbackHint
        ? `；建议：${options.fallbackHint}`
        : '';
      errorMessage.value = `${taskName} 失败：${message}（追踪ID: ${traceId}）${fallbackText}`;
      await logBridge('error', `${taskName}:fail`, traceId, {
        message,
        fallbackHint: options?.fallbackHint || '',
      });

      return null;
    }
  }

  return { runTask,
  };
}
