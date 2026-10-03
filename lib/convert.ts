import mammoth from 'mammoth/mammoth.browser.js';
import JSZip from 'jszip';
import { htmlToBlocks, IMAGE_SRC_PREFIX, TITLE_CLASS } from './blocks';
import type { ImportImage, ImportPayload } from './types';

const STYLE_MAP = [
  `p[style-name='Title'] => h1.${TITLE_CLASS}:fresh`,
  "p[style-name='Subtitle'] => h2:fresh",
  "p[style-name='Quote'] => blockquote > p:fresh",
  "p[style-name='Intense Quote'] => blockquote > p:fresh",
  "p[style-name='Block Text'] => blockquote > p:fresh",
  "p[style-name='Code'] => pre:separator('\\n')",
  "p[style-name='HTML Preformatted'] => pre:separator('\\n')",
];

export interface ConvertResult {
  payload: ImportPayload;
  /** Human-readable notes about content that couldn't be carried over exactly. */
  warnings: string[];
}

export async function convertDocx(data: ArrayBuffer, fileName: string): Promise<ConvertResult> {
  const images: ImportImage[] = [];
  const result = await mammoth.convertToHtml(
    { arrayBuffer: data },
    {
      styleMap: STYLE_MAP,
      convertImage: mammoth.images.imgElement(async (image) => {
        const index = images.length;
        const mime = image.contentType || 'application/octet-stream';
        images.push({ name: `image-${index + 1}.${extensionFor(mime)}`, mime, data: await image.readAsArrayBuffer() });
        return { src: `${IMAGE_SRC_PREFIX}${index}` };
      }),
    },
  );

  const { blocks, title, notes } = htmlToBlocks(result.value);
  const warnings: string[] = [];
  if (notes.tables) warnings.push(`${plural(notes.tables, 'table')} flattened to text (X Articles has no tables)`);
  if (notes.codeBlocks) warnings.push(`${plural(notes.codeBlocks, 'code block')} imported as plain text`);

  return {
    payload: {
      title: title ?? (await coreTitle(data)) ?? fileName.replace(/\.docx$/i, ''),
      blocks,
      images,
    },
    warnings,
  };
}

/** Reads dc:title from docProps/core.xml (File → Properties → Title in Word). */
async function coreTitle(data: ArrayBuffer): Promise<string | null> {
  try {
    const zip = await JSZip.loadAsync(data);
    const xml = await zip.file('docProps/core.xml')?.async('string');
    if (!xml) return null;
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const el = doc.getElementsByTagNameNS('http://purl.org/dc/elements/1.1/', 'title')[0];
    return el?.textContent?.trim() || null;
  } catch {
    return null;
  }
}

function extensionFor(mime: string): string {
  const sub = mime.split('/')[1] ?? 'bin';
  return { jpeg: 'jpg', 'svg+xml': 'svg', 'x-emf': 'emf', 'x-wmf': 'wmf' }[sub] ?? sub;
}

export function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? '' : 's'}`;
}
