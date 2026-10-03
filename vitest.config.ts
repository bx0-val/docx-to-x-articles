import { defineConfig } from 'vitest/config';
import { WxtVitest } from 'wxt/testing/vitest-plugin';

export default defineConfig({
  plugins: [WxtVitest()],
  // jsdom's ArrayBuffer is a different realm from Node's, which breaks
  // mammoth's zip reader, so run in Node and borrow only jsdom's DOMParser.
  test: { environment: 'node', setupFiles: ['tests/setup.ts'] },
});
