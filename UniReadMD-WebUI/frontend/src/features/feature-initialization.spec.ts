import { createPinia } from 'pinia';
import { flushPromises, mount } from '@vue/test-utils';
import { describe, expect, it } from 'vitest';
import App from '../App.vue';

describe('前端特征初始化', () => {
  it(
    'Given bridge 缺失，When 初始化工作台，Then 返回可观察的类型化错误',
    async () => {
      Object.defineProperty(window, 'nativeBridge', {
        configurable: true,
        value: undefined,
        writable: true,
      });

      const wrapper = mount(App, {
        global: {
          plugins: [createPinia()],
        },
      });
      await flushPromises();

      expect(wrapper.get('.error-banner').text()).toContain(
        'ERR_FEATURE_BRIDGE_UNAVAILABLE',
      );
      wrapper.unmount();
    },
  );
});
