import type { Tab } from './workspace-contracts';
import { ref } from 'vue';

export function useSidebarTabs() {
  const activeTab = ref<Tab>('outline');

  return { activeTab,
  };
}
