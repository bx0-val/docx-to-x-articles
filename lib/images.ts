import type { Block, ImportPayload } from './types';

// X accepts these directly; anything else is re-encoded to PNG if the browser can decode it.
const UPLOADABLE = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/webp']);

/**
 * Makes every image uploadable, dropping the ones the browser can't decode
 * (typically EMF/WMF vector graphics from older Office documents).
 */
export async function prepareImages(payload: ImportPayload): Promise<{ payload: ImportPayload; skipped: number }> {
  const images: ImportPayload['images'] = [];
  const remap = new Map<number, number>();
  let skipped = 0;

  for (const [index, image] of payload.images.entries()) {
    if (UPLOADABLE.has(image.mime)) {
      remap.set(index, images.push(image) - 1);
      continue;
    }
    const png = await toPng(image.data, image.mime);
    if (png) remap.set(index, images.push({ name: image.name.replace(/\.\w+$/, '.png'), mime: 'image/png', data: png }) - 1);
    else skipped++;
  }

  const blocks: Block[] = payload.blocks.flatMap((b): Block[] => {
    if (b.kind !== 'image') return [b];
    const next = remap.get(b.image);
    return next === undefined ? [] : [{ kind: 'image', image: next }];
  });

  return { payload: { ...payload, blocks, images }, skipped };
}

async function toPng(data: ArrayBuffer, mime: string): Promise<ArrayBuffer | null> {
  try {
    const bitmap = await createImageBitmap(new Blob([data], { type: mime }));
    const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
    canvas.getContext('2d')!.drawImage(bitmap, 0, 0);
    bitmap.close();
    return await (await canvas.convertToBlob({ type: 'image/png' })).arrayBuffer();
  } catch {
    return null;
  }
}
