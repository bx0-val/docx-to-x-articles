// Runs in the page's main world. X's article body is a Draft.js editor; we
// reach its EditorState and onChange through React's fiber tree and write
// blocks directly. Synthetic paste events lose formatting and can't carry
// images, which is why the body isn't pasted. Images go through X's own
// upload handler (`onFilesAdded`) so they're stored on X's media servers.
// Approach adapted from xPoster (MIT): https://github.com/nevertoday/xposter

import type { ImportPayload, ImportSummary, TextBlock } from './types';

// Draft.js / Immutable.js objects come from X's bundle and are untyped.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

interface DraftNode {
  props: { editorState: Any; onChange: (state: Any) => void };
}

const EDITOR_SELECTOR =
  ".public-DraftEditor-content[contenteditable='true'], [contenteditable='true'][role='textbox']";
const MARKER_PREFIX = '⟦docx2x-image-';
const UPLOAD_START_TIMEOUT_MS = 45_000;
const UPLOAD_FINISH_TIMEOUT_MS = 120_000;
/** X sometimes never fills in the media id client-side; accept a stable entity after this. */
const UPLOAD_PENDING_ACCEPT_MS = 25_000;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

export class ImportCancelled extends Error {}

export interface RunOptions {
  progress: (text: string) => void;
  /** Called when the draft already has content; resolve false to abort. */
  confirmReplace: () => boolean;
}

export async function runImport(payload: ImportPayload, opts: RunOptions): Promise<ImportSummary> {
  const summary: ImportSummary = { imagesUploaded: 0, imagesFailed: 0, titleSet: false };

  let node = findDraftNode();
  if (!node) throw new Error("Couldn't find X's article editor on this page. Click into the article body once, then drop the file again.");

  if (hasContent(node) && !opts.confirmReplace()) throw new ImportCancelled();

  opts.progress('Setting title…');
  summary.titleSet = await setTitle(payload.title);

  opts.progress('Writing text…');
  node = findDraftNode() ?? node;
  await writeBlocks(node, payload);

  const imageBlocks = payload.blocks.filter((b) => b.kind === 'image');
  for (const [i, block] of imageBlocks.entries()) {
    opts.progress(`Uploading image ${i + 1} of ${imageBlocks.length}…`);
    const image = payload.images[block.image];
    const marker = markerFor(block.image);
    const ok = image ? await uploadAtMarker(new File([image.data], image.name, { type: image.mime }), marker) : false;
    if (ok) summary.imagesUploaded++;
    else {
      summary.imagesFailed++;
      await replaceMarker(marker, `[Image ${i + 1} could not be uploaded]`);
    }
  }

  await removeStrayEmptyBlocks();
  // X's title field can lose a programmatic value if the editor re-renders first.
  const titleField = findTitleField();
  if (!summary.titleSet || (titleField && titleValue(titleField).trim() !== payload.title)) summary.titleSet = await setTitle(payload.title);
  return summary;
}

// ─── Finding the editor ────────────────────────────────────────────────────

/** The widest matching editable; an empty body editor can be very short. */
function findEditorElement(): HTMLElement | null {
  let best: HTMLElement | null = null;
  let bestWidth = 200;
  for (const el of document.querySelectorAll<HTMLElement>(EDITOR_SELECTOR)) {
    const { width, height } = el.getBoundingClientRect();
    if (height > 0 && width > bestWidth) [best, bestWidth] = [el, width];
  }
  return best;
}

function fiberOf(el: Element): Any {
  const key = Object.keys(el).find((k) => k.startsWith('__reactFiber$') || k.startsWith('__reactInternalInstance$'));
  return key ? (el as Any)[key] : null;
}

function findDraftNode(): DraftNode | null {
  const editor = findEditorElement();
  let fiber = editor && fiberOf(editor);
  for (let depth = 0; fiber && depth < 80; depth++) {
    const props = fiber.stateNode?.props;
    if (props?.editorState && typeof props.onChange === 'function') return fiber.stateNode;
    fiber = fiber.return;
  }
  return null;
}

function findOnFilesAdded(): ((files: File[]) => void) | null {
  const editor = findEditorElement();
  let fiber = editor && fiberOf(editor);
  for (let depth = 0; fiber && depth < 160; depth++) {
    const props = fiber.memoizedProps ?? fiber.stateNode?.props;
    if (typeof props?.onFilesAdded === 'function') return props.onFilesAdded;
    const nested = searchChildren(fiber.child, 0);
    if (nested) return nested;
    fiber = fiber.return;
  }
  return null;
}

function searchChildren(fiber: Any, depth: number): ((files: File[]) => void) | null {
  if (!fiber || depth > 8) return null;
  const props = fiber.memoizedProps ?? fiber.stateNode?.props;
  if (typeof props?.onFilesAdded === 'function') return props.onFilesAdded;
  return searchChildren(fiber.child, depth + 1) ?? searchChildren(fiber.sibling, depth);
}

function content(): Any {
  return findDraftNode()?.props.editorState.getCurrentContent();
}

function hasContent(node: DraftNode): boolean {
  let found = false;
  node.props.editorState
    .getCurrentContent()
    .getBlockMap()
    .forEach((block: Any) => {
      if (block.getType() === 'atomic' || block.getText().trim()) found = true;
    });
  return found;
}

/**
 * Hands X a new EditorState and waits for React to re-render with it. Reading
 * props.editorState before then returns the old state, and building on that
 * would silently undo this change.
 */
async function apply(node: DraftNode, next: Any): Promise<void> {
  const previous = node.props.editorState;
  node.props.onChange(next);
  for (let waited = 0; waited < 2000; waited += 20) {
    if (findDraftNode()?.props.editorState !== previous) return;
    await sleep(20);
  }
}

// ─── Title ─────────────────────────────────────────────────────────────────

type TitleField = HTMLInputElement | HTMLTextAreaElement | HTMLElement;

function findTitleField(): TitleField | null {
  const editor = findEditorElement();
  let best: TitleField | null = null;
  let bestScore = 0;
  const candidates = document.querySelectorAll<TitleField>("textarea, input[type='text'], input:not([type]), [contenteditable='true']");
  for (const el of candidates) {
    if (editor && (editor.contains(el) || el.contains(editor))) continue;
    const rect = el.getBoundingClientRect();
    if (rect.width < 4 || rect.height < 4) continue;
    const label = ['placeholder', 'aria-label', 'data-testid', 'name', 'data-placeholder']
      .map((a) => el.getAttribute(a))
      .join(' ')
      .toLowerCase();
    let score = 0;
    if (label.includes('title')) score += 10;
    if (label.includes('search')) score -= 20;
    if (rect.width > 240) score += 2;
    if (editor && rect.bottom <= editor.getBoundingClientRect().top + 4) score += 3;
    if (score > bestScore) [best, bestScore] = [el, score];
  }
  return bestScore >= 10 ? best : null;
}

function titleValue(field: TitleField): string {
  return field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement ? field.value : (field.textContent ?? '');
}

async function setTitle(title: string): Promise<boolean> {
  if (!title) return true;
  const field = findTitleField();
  if (!field) return false;
  field.focus();
  if (field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement) {
    const proto = field instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    // Bypass React's value tracker so its onChange sees the new value.
    Object.getOwnPropertyDescriptor(proto, 'value')?.set?.call(field, title);
    field.dispatchEvent(new Event('input', { bubbles: true }));
    field.dispatchEvent(new Event('change', { bubbles: true }));
  } else {
    document.execCommand('selectAll', false);
    document.execCommand('insertText', false, title);
  }
  field.blur();
  await sleep(100);
  return titleValue(field).trim() === title;
}

// ─── Writing text blocks ───────────────────────────────────────────────────

const markerFor = (image: number) => `${MARKER_PREFIX}${image}⟧`;

/**
 * Draft's constructors aren't exported, so clone a block and character from
 * the editor. An empty editor has none, but ContentState.createFromText
 * builds them with the same classes and settings X's editor uses.
 */
function characterSample(node: DraftNode): { block: Any; character: Any } {
  const contentState = node.props.editorState.getCurrentContent();
  let sample: { block: Any; character: Any } | null = null;
  contentState.getBlockMap().forEach((block: Any) => {
    const character = block.getCharacterList().first();
    if (character?.getStyle) {
      sample = { block, character };
      return false;
    }
  });
  if (sample) return sample;
  const block = contentState.constructor.createFromText('x').getFirstBlock();
  return { block, character: block.getCharacterList().first() };
}

async function writeBlocks(node: DraftNode, payload: ImportPayload) {
  const editorState = node.props.editorState;
  const EditorState = editorState.constructor;
  const SelectionState = editorState.getSelection().constructor;
  let contentState = editorState.getCurrentContent();
  const sample = characterSample(node);
  const BlockMap = contentState.getBlockMap().constructor;
  const CharacterList = sample.block.getCharacterList().constructor;
  const emptyStyle = sample.character.getStyle().clear();
  const emptyData = sample.block.getData().clear();

  let blockMap = BlockMap();
  let lastKey = '';
  payload.blocks.forEach((block, index) => {
    const text: TextBlock =
      block.kind === 'image'
        ? { kind: 'text', type: 'unstyled', text: markerFor(block.image), depth: 0, styles: [], links: [] }
        : block;

    const entities: { from: number; to: number; key: string }[] = [];
    for (const link of text.links) {
      contentState = contentState.createEntity('LINK', 'MUTABLE', { url: link.url });
      entities.push({ from: link.offset, to: link.offset + link.length, key: contentState.getLastCreatedEntityKey() });
    }

    let characters = CharacterList();
    for (let i = 0; i < text.text.length; i++) {
      let style = emptyStyle;
      for (const r of text.styles) if (i >= r.offset && i < r.offset + r.length) style = style.add(r.style);
      const entity = entities.find((e) => i >= e.from && i < e.to)?.key ?? null;
      characters = characters.push(sample.character.set('style', style).set('entity', entity));
    }

    const key = `d2x${index.toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    blockMap = blockMap.set(
      key,
      sample.block.merge({ key, type: text.type, text: text.text, characterList: characters, depth: text.depth, data: emptyData }),
    );
    lastKey = key;
  });

  if (!lastKey) return;
  const selection = SelectionState.createEmpty(lastKey);
  const next = contentState.set('blockMap', blockMap).set('selectionBefore', selection).set('selectionAfter', selection);
  await apply(node, EditorState.moveSelectionToEnd(EditorState.push(editorState, next, 'insert-fragment')));
}

// ─── Images ────────────────────────────────────────────────────────────────

function markerBlockKey(contentState: Any, marker: string): string | null {
  let found: string | null = null;
  contentState.getBlockMap().forEach((block: Any, key: string) => {
    if (!found && block.getType() !== 'atomic' && block.getText().includes(marker)) found = key;
  });
  return found;
}

interface MediaRef {
  blockKey: string;
  entityKey: string;
  ready: boolean;
}

function mediaBlocks(contentState: Any): MediaRef[] {
  const refs: MediaRef[] = [];
  contentState.getBlockMap().forEach((block: Any, blockKey: string) => {
    if (block.getType() !== 'atomic') return;
    const entityKey = block.getEntityAt(0);
    if (!entityKey) return;
    const entity = contentState.getEntity(entityKey);
    if (entity.getType() !== 'MEDIA') return;
    refs.push({ blockKey, entityKey, ready: hasMediaId(entity.getData()) });
  });
  return refs;
}

/** X fills in a media id once the upload is processed; its exact location varies. */
function hasMediaId(data: Any, depth = 0): boolean {
  if (!data || typeof data !== 'object' || depth > 4) return false;
  for (const [k, v] of Object.entries(data)) {
    if (/media_?id/i.test(k) && v) return true;
    if (typeof v === 'object' && hasMediaId(v, depth + 1)) return true;
  }
  return false;
}

async function uploadAtMarker(file: File, marker: string): Promise<boolean> {
  const node = findDraftNode();
  const onFilesAdded = findOnFilesAdded();
  if (!node || !onFilesAdded) return false;

  const editorState = node.props.editorState;
  const markerKey = markerBlockKey(editorState.getCurrentContent(), marker);
  if (!markerKey) return false;
  const before = new Set(mediaBlocks(editorState.getCurrentContent()).map((m) => m.entityKey));

  // Put the caret on the placeholder so X inserts the image there.
  const SelectionState = editorState.getSelection().constructor;
  await apply(node, editorState.constructor.forceSelection(editorState, SelectionState.createEmpty(markerKey)));
  try {
    onFilesAdded([file]);
  } catch (e) {
    console.warn('[docx2x] upload handler failed', e);
    return false;
  }

  const start = Date.now();
  let firstSeen = 0;
  while (Date.now() - start < UPLOAD_FINISH_TIMEOUT_MS) {
    await sleep(350);
    const cs = content();
    if (!cs) continue;
    const added = mediaBlocks(cs).find((m) => !before.has(m.entityKey));
    if (!added) {
      if (Date.now() - start > UPLOAD_START_TIMEOUT_MS) return false;
      continue;
    }
    firstSeen ||= Date.now();
    if (added.ready || Date.now() - firstSeen > UPLOAD_PENDING_ACCEPT_MS) {
      await placeAtMarker(added.blockKey, marker);
      return true;
    }
  }
  return false;
}

/** Moves the uploaded image block to where its placeholder is, then deletes the placeholder. */
async function placeAtMarker(imageKey: string, marker: string) {
  const node = findDraftNode();
  if (!node) return;
  const editorState = node.props.editorState;
  const contentState = editorState.getCurrentContent();
  const markerKey = markerBlockKey(contentState, marker);
  const blockMap = contentState.getBlockMap();
  let next = blockMap.constructor();
  blockMap.forEach((block: Any, key: string) => {
    if (key === imageKey) return;
    if (key === markerKey) next = next.set(imageKey, blockMap.get(imageKey));
    else next = next.set(key, block);
  });
  if (!markerKey) return; // Placeholder vanished; leave the image where X put it.
  await pushBlockMap(node, next, 'remove-range');
}

async function replaceMarker(marker: string, text: string) {
  const node = findDraftNode();
  if (!node) return;
  const contentState = node.props.editorState.getCurrentContent();
  const key = markerBlockKey(contentState, marker);
  if (!key) return;
  const block = contentState.getBlockForKey(key);
  const character = block.getCharacterList().first();
  const characters = block.getCharacterList().clear().concat(Array.from({ length: text.length }, () => character));
  await pushBlockMap(node, contentState.getBlockMap().set(key, block.merge({ text, characterList: characters })), 'change-block-data');
}

/**
 * Draft splits the placeholder paragraph around each inserted image, which can
 * leave empty paragraphs behind. The import never writes empty paragraphs, so
 * any left are strays. Keep the last block so the caret has a home.
 */
async function removeStrayEmptyBlocks() {
  const node = findDraftNode();
  if (!node) return;
  const blockMap = node.props.editorState.getCurrentContent().getBlockMap();
  const lastKey = blockMap.last()?.getKey();
  const next = blockMap.filter(
    (block: Any, key: string) => key === lastKey || block.getType() === 'atomic' || block.getText().trim() !== '',
  );
  if (next.size !== blockMap.size) await pushBlockMap(node, next, 'remove-range');
}

async function pushBlockMap(node: DraftNode, blockMap: Any, changeType: string) {
  const editorState = node.props.editorState;
  const EditorState = editorState.constructor;
  const SelectionState = editorState.getSelection().constructor;
  const selection = SelectionState.createEmpty(blockMap.last().getKey());
  const next = editorState
    .getCurrentContent()
    .set('blockMap', blockMap)
    .set('selectionBefore', selection)
    .set('selectionAfter', selection);
  await apply(node, EditorState.moveSelectionToEnd(EditorState.push(editorState, next, changeType)));
}
