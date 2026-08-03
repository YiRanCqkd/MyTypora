import type { ExportCapabilitiesResult } from '../../types/native';
import { ref } from 'vue';

export function useExportCapabilitiesState() {
  const exportCapabilities = ref<ExportCapabilitiesResult | null>(null);

  return { exportCapabilities,
  };
}
