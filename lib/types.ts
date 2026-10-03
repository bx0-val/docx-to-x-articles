// Shapes shared between the isolated content script (which parses the .docx)
// and the main-world bridge (which writes into X's Draft.js editor).

export type InlineStyle = 'BOLD' | 'ITALIC' | 'STRIKETHROUGH';

export type TextBlockType =
  | 'unstyled'
  | 'header-one'
  | 'header-two'
  | 'blockquote'
  | 'unordered-list-item'
  | 'ordered-list-item';

export interface StyleRange {
  offset: number;
  length: number;
  style: InlineStyle;
}

export interface LinkRange {
  offset: number;
  length: number;
  url: string;
}

export interface TextBlock {
  kind: 'text';
  type: TextBlockType;
  text: string;
  depth: number;
  styles: StyleRange[];
  links: LinkRange[];
}

export interface ImageBlock {
  kind: 'image';
  /** Index into ImportPayload.images. */
  image: number;
}

export type Block = TextBlock | ImageBlock;

export interface ImportImage {
  name: string;
  mime: string;
  data: ArrayBuffer;
}

export interface ImportPayload {
  title: string;
  blocks: Block[];
  images: ImportImage[];
}

export interface ImportSummary {
  imagesUploaded: number;
  imagesFailed: number;
  titleSet: boolean;
}

// window.postMessage protocol between the two worlds.
export const TO_MAIN = 'docx2x';
export const FROM_MAIN = 'docx2x-main';

export type ToMainMessage = { source: typeof TO_MAIN; kind: 'run'; id: string; payload: ImportPayload };

export type FromMainMessage =
  | { source: typeof FROM_MAIN; kind: 'progress'; id: string; text: string }
  | { source: typeof FROM_MAIN; kind: 'done'; id: string; summary: ImportSummary }
  | { source: typeof FROM_MAIN; kind: 'cancelled'; id: string }
  | { source: typeof FROM_MAIN; kind: 'error'; id: string; error: string };
