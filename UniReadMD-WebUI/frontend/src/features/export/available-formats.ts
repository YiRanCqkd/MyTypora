import type { ExportFormat } from '../../types/native';
import { computed } from 'vue';
import { BUILTIN_EXPORT_FORMATS } from './export-formats';
import type { useExportCapabilitiesState } from './export-capabilities-state';
type UseAvailableExportFormatsDependencies =
  ReturnType<typeof useExportCapabilitiesState>;

export function useAvailableExportFormats(input: UseAvailableExportFormatsDependencies) {
  const { exportCapabilities,
  } = input;

  const availableExportFormats = computed(() => {
    if (!exportCapabilities.value) {
      return new Set<ExportFormat>(BUILTIN_EXPORT_FORMATS);
    }

    const formats = [
      ...exportCapabilities.value.builtinFormats,
      ...exportCapabilities.value.pandocFormats,
    ];
    const uniqueFormats = formats.filter((item, index) => formats.indexOf(item) === index);
    return new Set<ExportFormat>(uniqueFormats);
  });

  return { availableExportFormats,
  };
}
