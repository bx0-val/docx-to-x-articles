import type { Block, InlineStyle, LinkRange, StyleRange, TextBlock, TextBlockType } from './types';

// Turns mammoth's HTML into the flat block list X's Draft.js editor understands.
// X Articles supports two heading levels, quotes, bulleted/numbered lists,
// bold/italic/strikethrough, links and images; everything else is downgraded.

export const IMAGE_SRC_PREFIX = 'docx-image:';
export const TITLE_CLASS = 'docx-title';
const MAX_LIST_DEPTH = 4;

export interface ConvertNotes {
  tables: number;
  codeBlocks: number;
}

export interface HtmlToBlocksResult {
  blocks: Block[];
  /** Text of a paragraph styled as Word's "Title", or a leading Heading 1. */
  title: string | null;
  notes: ConvertNotes;
}

interface Builder {
  text: string;
  styles: StyleRange[];
  links: LinkRange[];
}

interface InlineContext {
  styles: InlineStyle[];
  url: string | null;
}

export function htmlToBlocks(html: string): HtmlToBlocksResult {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const blocks: Block[] = [];
  const notes: ConvertNotes = { tables: 0, codeBlocks: 0 };
  let title: string | null = null;

  const titleEl = doc.body.querySelector(`h1.${TITLE_CLASS}`);
  if (titleEl) {
    title = titleEl.textContent?.trim() || null;
    titleEl.remove();
  }

  const emitInline = (container: Node, type: TextBlockType, depth = 0, skip?: (n: Node) => boolean) => {
    let current = newBuilder();
    const flush = () => {
      const block = finishBlock(current, type, depth);
      if (block) blocks.push(block);
      current = newBuilder();
    };
    const walk = (node: Node, ctx: InlineContext) => {
      if (skip?.(node)) return;
      if (node.nodeType === Node.TEXT_NODE) {
        append(current, node.textContent ?? '', ctx);
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const el = node as Element;
      const tag = el.tagName.toLowerCase();
      if (tag === 'img') {
        const index = imageIndex(el);
        if (index !== null) {
          flush();
          blocks.push({ kind: 'image', image: index });
        }
        return;
      }
      if (tag === 'br') {
        append(current, '\n', ctx);
        return;
      }
      // Separate multiple paragraphs inside one list item / quote with a soft break.
      if (tag === 'p' && current.text.trim()) append(current, '\n', ctx);
      const next: InlineContext = { styles: ctx.styles, url: ctx.url };
      if (tag === 'strong' || tag === 'b') next.styles = [...ctx.styles, 'BOLD'];
      else if (tag === 'em' || tag === 'i') next.styles = [...ctx.styles, 'ITALIC'];
      else if (tag === 's' || tag === 'del' || tag === 'strike') next.styles = [...ctx.styles, 'STRIKETHROUGH'];
      else if (tag === 'a') next.url = safeUrl(el.getAttribute('href'));
      // Footnotes keep their "[1]" markers but drop the "↑" back-links.
      if (tag === 'a' && /^#(footnote|endnote)-ref-/.test(el.getAttribute('href') ?? '')) return;
      for (const child of Array.from(el.childNodes)) walk(child, next);
    };
    for (const child of Array.from(container.childNodes)) walk(child, { styles: [], url: null });
    flush();
  };

  const emitList = (list: Element, depth: number) => {
    const type: TextBlockType = list.tagName.toLowerCase() === 'ol' ? 'ordered-list-item' : 'unordered-list-item';
    for (const li of Array.from(list.children)) {
      if (li.tagName.toLowerCase() !== 'li') continue;
      const isList = (n: Node) => n instanceof Element && /^(ul|ol)$/i.test(n.tagName);
      // mammoth can wrap list item content in <p> when the item has multiple paragraphs.
      emitInline(li, type, Math.min(depth, MAX_LIST_DEPTH), isList);
      for (const nested of Array.from(li.children)) {
        if (isList(nested)) emitList(nested, depth + 1);
      }
    }
  };

  const emitBlock = (el: Element) => {
    const tag = el.tagName.toLowerCase();
    switch (tag) {
      case 'h1':
        if (title === null && blocks.length === 0) {
          title = el.textContent?.trim() || null;
          if (title) return;
        }
        return emitInline(el, 'header-one');
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        return emitInline(el, 'header-two');
      case 'blockquote': {
        const paragraphs = Array.from(el.children).filter((c) => c.tagName.toLowerCase() === 'p');
        if (paragraphs.length) paragraphs.forEach((p) => emitInline(p, 'blockquote'));
        else emitInline(el, 'blockquote');
        return;
      }
      case 'ul':
      case 'ol':
        return emitList(el, 0);
      case 'table':
        notes.tables++;
        for (const row of Array.from(el.querySelectorAll('tr'))) {
          const cells = Array.from(row.children).filter((c) => /^t[dh]$/i.test(c.tagName));
          const holder = el.ownerDocument.createElement('p');
          cells.forEach((cell, i) => {
            if (i > 0) holder.append(' | ');
            // Cells hold <p>s; inline their contents so the row becomes one line.
            for (const child of Array.from(cell.childNodes)) {
              if (child instanceof Element && child.tagName.toLowerCase() === 'p') holder.append(...Array.from(child.childNodes));
              else holder.append(child);
            }
          });
          emitInline(holder, 'unstyled');
        }
        return;
      case 'pre':
        notes.codeBlocks++;
        return emitInline(el, 'unstyled');
      default:
        return emitInline(el, 'unstyled');
    }
  };

  for (const el of Array.from(doc.body.children)) emitBlock(el);

  return { blocks, title, notes };
}

function newBuilder(): Builder {
  return { text: '', styles: [], links: [] };
}

function append(b: Builder, raw: string, ctx: InlineContext) {
  const text = raw.replace(/[\t\r]/g, ' ');
  if (!text) return;
  const offset = b.text.length;
  b.text += text;
  for (const style of new Set(ctx.styles)) b.styles.push({ offset, length: text.length, style });
  if (ctx.url) b.links.push({ offset, length: text.length, url: ctx.url });
}

/** Trims the block, shifts/merges ranges, and drops it if empty. */
function finishBlock(b: Builder, type: TextBlockType, depth: number): TextBlock | null {
  const start = b.text.length - b.text.trimStart().length;
  const text = b.text.trim();
  if (!text) return null;
  const end = start + text.length;
  const clip = <T extends { offset: number; length: number }>(r: T): T | null => {
    const from = Math.max(r.offset, start);
    const to = Math.min(r.offset + r.length, end);
    return to > from ? { ...r, offset: from - start, length: to - from } : null;
  };
  const styles = mergeRanges(b.styles.map(clip).filter((r): r is StyleRange => r !== null), (r) => r.style);
  const links = mergeRanges(b.links.map(clip).filter((r): r is LinkRange => r !== null), (r) => r.url);
  return { kind: 'text', type, text, depth: type.endsWith('list-item') ? depth : 0, styles, links };
}

/** Joins adjacent ranges with the same key (mammoth often splits runs). */
function mergeRanges<T extends { offset: number; length: number }>(ranges: T[], key: (r: T) => string): T[] {
  const sorted = [...ranges].sort((a, b) => key(a).localeCompare(key(b)) || a.offset - b.offset);
  const out: T[] = [];
  for (const r of sorted) {
    const last = out[out.length - 1];
    if (last && key(last) === key(r) && last.offset + last.length >= r.offset) {
      last.length = Math.max(last.offset + last.length, r.offset + r.length) - last.offset;
    } else {
      out.push({ ...r });
    }
  }
  return out.sort((a, b) => a.offset - b.offset);
}

function imageIndex(img: Element): number | null {
  const src = img.getAttribute('src') ?? '';
  if (!src.startsWith(IMAGE_SRC_PREFIX)) return null;
  const n = Number(src.slice(IMAGE_SRC_PREFIX.length));
  return Number.isInteger(n) ? n : null;
}

function safeUrl(href: string | null): string | null {
  if (!href) return null;
  try {
    const url = new URL(href);
    return ['http:', 'https:', 'mailto:'].includes(url.protocol) ? url.href : null;
  } catch {
    return null;
  }
}
