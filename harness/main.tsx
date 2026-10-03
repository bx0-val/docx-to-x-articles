// A stand-in for X's article editor: a Draft.js editor with a title textarea
// and an `onFilesAdded` prop that "uploads" by inserting an atomic MEDIA
// block at the selection, then fills in a media id later like X does.
import { AtomicBlockUtils, convertToRaw, Editor, EditorState, type ContentBlock } from 'draft-js';
import React, { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import 'draft-js/dist/Draft.css';
import { convertDocx } from '../lib/convert';
import { runImport } from '../lib/draft';

declare global {
  interface Window {
    harness: {
      importDocx(base64: string, opts?: { confirm?: boolean }): Promise<unknown>;
      raw(): unknown;
      title(): string;
      typeInEditor(text: string): void;
    };
  }
}

let latest: EditorState = EditorState.createEmpty();

function Media({ block, contentState }: { block: ContentBlock; contentState: any }) {
  const data = contentState.getEntity(block.getEntityAt(0)).getData();
  return <figure>{data.mediaId ? `media ${data.mediaId}` : 'uploading…'}</figure>;
}

function MediaDropZone(props: { onFilesAdded: (files: File[]) => void; children: React.ReactNode }) {
  return <section>{props.children}</section>;
}

function App() {
  const [editorState, setEditorState] = useState(() => EditorState.createEmpty());
  const [title, setTitle] = useState('');
  const stateRef = useRef(editorState);
  const update = (s: EditorState) => {
    stateRef.current = latest = s;
    setEditorState(s);
  };

  const onFilesAdded = (files: File[]) => {
    for (const file of files) {
      let s = stateRef.current;
      const content = s.getCurrentContent().createEntity('MEDIA', 'IMMUTABLE', { name: file.name, size: file.size });
      const entityKey = content.getLastCreatedEntityKey();
      s = AtomicBlockUtils.insertAtomicBlock(EditorState.set(s, { currentContent: content }), entityKey, ' ');
      update(s);
      // Finish the "upload" later, the way X fills in media ids asynchronously.
      setTimeout(() => {
        const cur = stateRef.current;
        const next = cur.getCurrentContent().mergeEntityData(entityKey, { media_id: `m${Math.floor(Math.random() * 1e6)}` });
        update(EditorState.push(cur, next, 'apply-entity'));
      }, 600);
    }
  };

  return (
    <main>
      <textarea placeholder="Add a title" value={title} onChange={(e) => setTitle(e.target.value)} rows={1} />
      <MediaDropZone onFilesAdded={onFilesAdded}>
        <Editor
          editorState={editorState}
          onChange={update}
          blockRendererFn={(b) => (b.getType() === 'atomic' ? { component: Media, editable: false } : null)}
        />
      </MediaDropZone>
    </main>
  );
}

createRoot(document.getElementById('root')!).render(<App />);

window.harness = {
  async importDocx(base64, opts = {}) {
    const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
    const { payload } = await convertDocx(bytes.buffer, 'sample.docx');
    const progress: string[] = [];
    const summary = await runImport(payload, { progress: (t) => progress.push(t), confirmReplace: () => opts.confirm ?? true });
    return { summary, progress };
  },
  raw: () => {
    const content = latest.getCurrentContent();
    const raw = convertToRaw(content);
    return raw.blocks.map((b) => ({
      type: b.type,
      text: b.text,
      depth: b.depth,
      styles: b.inlineStyleRanges,
      entity: b.entityRanges.map((r) => {
        const e = raw.entityMap[r.key];
        return { type: e?.type, data: e?.data, offset: r.offset, length: r.length };
      }),
    }));
  },
  title: () => document.querySelector('textarea')!.value,
  typeInEditor: (text) => {
    (document.querySelector('[contenteditable=true]') as HTMLElement).focus();
    document.execCommand('insertText', false, text);
  },
};
