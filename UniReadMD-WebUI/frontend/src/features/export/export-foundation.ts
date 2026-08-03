import { useAvailableExportFormats } from './available-formats';
import { useExportCapabilitiesState } from './export-capabilities-state';

export function useExportFoundation() {
  const capabilities = useExportCapabilitiesState();
  const formats = useAvailableExportFormats(capabilities);

  return {
    ...capabilities,
    ...formats,
  };
}
