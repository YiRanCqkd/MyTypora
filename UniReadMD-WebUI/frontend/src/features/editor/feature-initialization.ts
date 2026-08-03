export const FEATURE_BRIDGE_UNAVAILABLE = 'ERR_FEATURE_BRIDGE_UNAVAILABLE';

export class FeatureInitializationError extends Error {
  readonly code: typeof FEATURE_BRIDGE_UNAVAILABLE;

  constructor() {
    super('nativeBridge 尚未完成初始化');
    this.name = 'FeatureInitializationError';
    this.code = FEATURE_BRIDGE_UNAVAILABLE;
  }

  toDisplayMessage() {
    return `${this.code}: ${this.message}`;
  }
}

export function resolveFeatureInitializationError(
  nativeBridge: Window['nativeBridge'] | undefined,
) {
  return nativeBridge ? null : new FeatureInitializationError();
}
