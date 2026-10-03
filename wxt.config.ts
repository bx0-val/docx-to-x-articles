import { defineConfig } from 'wxt';

export default defineConfig({
  // MV3 for both browsers: Firefox needs it (128+) for main-world content scripts.
  manifestVersion: 3,
  manifest: ({ browser }) => ({
    name: 'Docx to X Articles',
    description: 'Drag a Word (.docx) file onto an X Articles draft to import its title, formatting and images.',
    browser_specific_settings: browser === 'firefox' ? {
      gecko: {
        id: 'docx-to-x-articles@extension',
        strict_min_version: '128.0',
        // Required by AMO for new extensions; nothing leaves the browser.
        data_collection_permissions: { required: ['none'] },
      },
    } as Record<string, unknown> : undefined,
  }),
});
