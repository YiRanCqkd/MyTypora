import { ref } from 'vue';

export function useApplicationStatus() {
  const appVersion = ref('0.0.0');

  const errorMessage = ref('');

  return { appVersion,errorMessage,
  };
}
