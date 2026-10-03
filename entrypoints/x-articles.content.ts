import { convertDocx, plural } from '@/lib/convert';
import { prepareImages } from '@/lib/images';
import { createOverlay, type Overlay } from '@/lib/overlay';
import { FROM_MAIN, TO_MAIN, type FromMainMessage, type ImportPayload, type ImportSummary, type ToMainMessage } from '@/lib/types';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const ARTICLE_PATH = /^\/compose\/articles?\/edit\//;

// Isolated-world half: handles the drag & drop, parses the .docx, and hands
// the result to editor-bridge.content.ts in the page's main world.
// X is a single-page app, so this runs on every x.com page and checks the
// path when a drag starts.
export default defineContentScript({
  matches: ['https://x.com/*', 'https://twitter.com/*'],
  runAt: 'document_idle',
  main() {
    let overlay: Overlay | null = null;
    const ui = () => (overlay ??= createOverlay());
    let busy = false;
    let dragDepth = 0;

    const onArticlePage = () => ARTICLE_PATH.test(location.pathname);

    // During a drag only MIME types are visible, not names. Some systems
    // (often Linux) report an empty type for .docx, so treat "unknown" as a
    // candidate; known non-docx types (e.g. images) are left to X.
    const isDocxDrag = (e: DragEvent) => {
      const items = Array.from(e.dataTransfer?.items ?? []).filter((i) => i.kind === 'file');
      return items.length > 0 && items.some((i) => i.type === DOCX_MIME || i.type === '');
    };

    const intercept = (e: DragEvent) => {
      if (!onArticlePage() || !isDocxDrag(e)) return false;
      e.preventDefault();
      e.stopPropagation();
      return true;
    };

    window.addEventListener(
      'dragenter',
      (e) => {
        if (!intercept(e)) return;
        dragDepth++;
        if (!busy) ui().showDropZone(true);
      },
      true,
    );
    window.addEventListener(
      'dragover',
      (e) => {
        if (intercept(e)) e.dataTransfer!.dropEffect = busy ? 'none' : 'copy';
      },
      true,
    );
    window.addEventListener(
      'dragleave',
      (e) => {
        if (!intercept(e)) return;
        if (--dragDepth <= 0) {
          dragDepth = 0;
          ui().showDropZone(false);
        }
      },
      true,
    );
    window.addEventListener(
      'drop',
      (e) => {
        if (!intercept(e)) return;
        dragDepth = 0;
        ui().showDropZone(false);
        if (busy) return;
        const file = Array.from(e.dataTransfer?.files ?? []).find((f) => f.type === DOCX_MIME || /\.docx$/i.test(f.name));
        if (!file) {
          ui().result('error', 'That isn’t a .docx file', ['Save the document as Word (.docx) and drop it again. Older .doc files aren’t supported.']);
          return;
        }
        busy = true;
        importFile(file, ui()).finally(() => (busy = false));
      },
      true,
    );
  },
});

async function importFile(file: File, ui: Overlay) {
  try {
    ui.busy(`Reading ${file.name}…`);
    const converted = await convertDocx(await file.arrayBuffer(), file.name);
    const { payload, skipped } = await prepareImages(converted.payload);
    const warnings = [...converted.warnings];
    if (skipped) warnings.push(`${plural(skipped, 'image')} skipped (format not supported by browsers, e.g. EMF/WMF)`);
    if (!payload.blocks.length) {
      ui.result('error', 'The document looks empty', ['No text or images were found in it.']);
      return;
    }

    const summary = await sendToEditor(payload, (text) => ui.busy(text));
    if (!summary) {
      ui.result('ok', 'Import cancelled');
      return;
    }
    if (!summary.titleSet) warnings.push('Couldn’t find the title field. Paste the title in yourself.');
    if (summary.imagesFailed) warnings.push(`${plural(summary.imagesFailed, 'image')} failed to upload and ${summary.imagesFailed === 1 ? 'was' : 'were'} marked in the text`);
    const images = summary.imagesUploaded ? ` with ${plural(summary.imagesUploaded, 'image')}` : '';
    ui.result('ok', `Imported “${payload.title}”${images}`, warnings);
  } catch (error) {
    console.error('[docx2x]', error);
    ui.result('error', 'Import failed', [error instanceof Error ? error.message : String(error)]);
  }
}

/** Resolves with the summary, or null if the user cancelled. */
function sendToEditor(payload: ImportPayload, progress: (text: string) => void): Promise<ImportSummary | null> {
  const id = crypto.randomUUID();
  return new Promise((resolve, reject) => {
    const onMessage = (event: MessageEvent<FromMainMessage>) => {
      const msg = event.data;
      if (event.source !== window || msg?.source !== FROM_MAIN || msg.id !== id) return;
      if (msg.kind === 'progress') return progress(msg.text);
      window.removeEventListener('message', onMessage);
      if (msg.kind === 'done') resolve(msg.summary);
      else if (msg.kind === 'cancelled') resolve(null);
      else reject(new Error(msg.error));
    };
    window.addEventListener('message', onMessage);
    const message: ToMainMessage = { source: TO_MAIN, kind: 'run', id, payload };
    window.postMessage(message, location.origin);
  });
}
