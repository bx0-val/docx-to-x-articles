import { ImportCancelled, runImport } from '@/lib/draft';
import { FROM_MAIN, TO_MAIN, type FromMainMessage, type ToMainMessage } from '@/lib/types';

type WithoutSource<T> = T extends unknown ? Omit<T, 'source'> : never;

// Main-world half: needs page JS objects (React fibers, Draft.js state) that
// the isolated content script can't see.
export default defineContentScript({
  matches: ['https://x.com/*', 'https://twitter.com/*'],
  world: 'MAIN',
  runAt: 'document_idle',
  main() {
    const post = (msg: WithoutSource<FromMainMessage>) => window.postMessage({ source: FROM_MAIN, ...msg }, location.origin);

    window.addEventListener('message', (event: MessageEvent<ToMainMessage>) => {
      if (event.source !== window || event.data?.source !== TO_MAIN || event.data.kind !== 'run') return;
      const { id, payload } = event.data;
      runImport(payload, {
        progress: (text) => post({ kind: 'progress', id, text }),
        confirmReplace: () => window.confirm('This article already has content. Replace it with the Word document?'),
      })
        .then((summary) => post({ kind: 'done', id, summary }))
        .catch((error: unknown) => {
          if (error instanceof ImportCancelled) return post({ kind: 'cancelled', id });
          console.error('[docx2x]', error);
          post({ kind: 'error', id, error: error instanceof Error ? error.message : String(error) });
        });
    });
  },
});
