import { createApp } from 'vue';
import { createPinia } from 'pinia';
import App from './App.vue';
import { installNativeBridgeFallback } from './services/native-bridge-fallback';
import './styles.css';

installNativeBridgeFallback();

Object.defineProperty(window, '__unireadmdBooted', {
  configurable: true,
  writable: true,
  value: true,
});

const app = createApp(App);

app.use(createPinia());
app.mount('#app');
