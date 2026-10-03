import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { htmlToBlocks } from '@/lib/blocks';
import { convertDocx } from '@/lib/convert';
import type { TextBlock } from '@/lib/types';

const fixture = () => {
  const buf = readFileSync(new URL('../fixtures/sample.docx', import.meta.url));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
};

const texts = (blocks: ReturnType<typeof htmlToBlocks>['blocks']) =>
  blocks.map((b) => (b.kind === 'image' ? `[img ${b.image}]` : `${b.type}${b.depth ? `@${b.depth}` : ''}: ${b.text}`));

describe('convertDocx (fixtures/sample.docx)', async () => {
  const { payload, warnings } = await convertDocx(fixture(), 'sample.docx');

  it('uses the Title-styled paragraph as the title and drops it from the body', () => {
    expect(payload.title).toBe('My Imported Article');
    expect(texts(payload.blocks)).not.toContain('header-one: My Imported Article');
  });

  it('maps structure to X block types in document order', () => {
    expect(texts(payload.blocks)).toEqual([
      'unstyled: Plain, bold, italic, struck, underlined and a link.',
      'header-one: Section heading',
      'header-two: Sub heading',
      'header-two: Third level',
      '[img 0]',
      'unordered-list-item: Bullet one',
      'unordered-list-item@1: Nested bullet',
      'unordered-list-item: Bullet two',
      'ordered-list-item: First step',
      'ordered-list-item: Second step',
      'blockquote: A wise quote.',
      'unstyled: Name | Value',
      'unstyled: a | 1',
      '[img 1]',
      'unstyled: The end.',
    ]);
  });

  it('keeps inline styles and links with correct offsets', () => {
    const p = payload.blocks[0] as TextBlock;
    const slice = (r: { offset: number; length: number }) => p.text.slice(r.offset, r.offset + r.length);
    expect(p.styles.map((r) => [r.style, slice(r)])).toEqual([
      ['BOLD', 'bold'],
      ['ITALIC', 'italic'],
      ['STRIKETHROUGH', 'struck'],
    ]);
    expect(p.links.map((l) => [slice(l), l.url])).toEqual([['link', 'https://example.com/']]);
  });

  it('extracts images as uploadable files', () => {
    expect(payload.images.map((i) => [i.name, i.mime])).toEqual([
      ['image-1.png', 'image/png'],
      ['image-2.png', 'image/png'],
    ]);
    expect(new Uint8Array(payload.images[0]!.data).slice(1, 4)).toEqual(new TextEncoder().encode('PNG'));
  });

  it('reports downgraded content', () => {
    expect(warnings).toEqual(['1 table flattened to text (X Articles has no tables)']);
  });
});

describe('htmlToBlocks', () => {
  it('uses a leading Heading 1 as the title when there is no Title style', () => {
    const { title, blocks } = htmlToBlocks('<h1>Hello</h1><p>Body</p><h1>Later</h1>');
    expect(title).toBe('Hello');
    expect(texts(blocks)).toEqual(['unstyled: Body', 'header-one: Later']);
  });

  it('leaves the title empty when the document starts with a paragraph', () => {
    expect(htmlToBlocks('<p>Intro</p><h1>Heading</h1>').title).toBeNull();
  });

  it('merges split runs, trims whitespace and clips ranges', () => {
    const [block] = htmlToBlocks('<p>  <strong>bo</strong><strong>ld</strong> text </p>').blocks as TextBlock[];
    expect(block!.text).toBe('bold text');
    expect(block!.styles).toEqual([{ offset: 0, length: 4, style: 'BOLD' }]);
  });

  it('splits a paragraph around an inline image', () => {
    const { blocks } = htmlToBlocks('<p>before<img src="docx-image:0">after</p>');
    expect(texts(blocks)).toEqual(['unstyled: before', '[img 0]', 'unstyled: after']);
  });

  it('drops unsafe links and footnote back-links but keeps footnote markers', () => {
    const { blocks } = htmlToBlocks(
      '<p>Claim<sup><a href="#footnote-1" id="footnote-ref-1">[1]</a></sup> <a href="javascript:alert(1)">x</a></p>' +
        '<ol><li id="footnote-1"><p>Source <a href="#footnote-ref-1">↑</a></p></li></ol>',
    );
    expect(texts(blocks)).toEqual(['unstyled: Claim[1] x', 'ordered-list-item: Source']);
    expect((blocks[0] as TextBlock).links).toEqual([]);
  });

  it('keeps line breaks and skips empty paragraphs', () => {
    expect(texts(htmlToBlocks('<p>one<br />two</p><p> </p><p></p>').blocks)).toEqual(['unstyled: one\ntwo']);
  });
});
