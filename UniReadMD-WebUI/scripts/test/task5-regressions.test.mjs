import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);

require('./packaging-flow.test.cjs');
require('./transient-rename.test.cjs');
