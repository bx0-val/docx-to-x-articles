// Generates fixtures/sample.docx covering everything the importer handles.
import { writeFileSync } from 'node:fs';
import {
  Document, ExternalHyperlink, HeadingLevel, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun,
} from 'docx';
import { encodePng } from './png.mjs';

const square = (r, g, b) => encodePng(64, 64, () => [r, g, b, 255]);
const image = (data) => new Paragraph({ children: [new ImageRun({ type: 'png', data, transformation: { width: 64, height: 64 } })] });
const cell = (text, bold = false) => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold })] })] });

const doc = new Document({
  title: 'Core Properties Title',
  styles: { paragraphStyles: [{ id: 'Quote', name: 'Quote', basedOn: 'Normal', run: { italics: true } }] },
  numbering: {
    config: [{ reference: 'numbers', levels: [{ level: 0, format: 'decimal', text: '%1.' }] }],
  },
  sections: [
    {
      children: [
        new Paragraph({ text: 'My Imported Article', heading: HeadingLevel.TITLE }),
        new Paragraph({
          children: [
            new TextRun('Plain, '),
            new TextRun({ text: 'bold', bold: true }),
            new TextRun(', '),
            new TextRun({ text: 'italic', italics: true }),
            new TextRun(', '),
            new TextRun({ text: 'struck', strike: true }),
            new TextRun(', '),
            new TextRun({ text: 'underlined', underline: {} }),
            new TextRun(' and a '),
            new ExternalHyperlink({ link: 'https://example.com/', children: [new TextRun({ text: 'link', style: 'Hyperlink' })] }),
            new TextRun('.'),
          ],
        }),
        new Paragraph({ text: 'Section heading', heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: 'Sub heading', heading: HeadingLevel.HEADING_2 }),
        new Paragraph({ text: 'Third level', heading: HeadingLevel.HEADING_3 }),
        image(square(29, 155, 240)),
        new Paragraph({ text: 'Bullet one', bullet: { level: 0 } }),
        new Paragraph({ text: 'Nested bullet', bullet: { level: 1 } }),
        new Paragraph({ text: 'Bullet two', bullet: { level: 0 } }),
        new Paragraph({ text: 'First step', numbering: { reference: 'numbers', level: 0 } }),
        new Paragraph({ text: 'Second step', numbering: { reference: 'numbers', level: 0 } }),
        new Paragraph({ text: 'A wise quote.', style: 'Quote' }),
        new Table({
          rows: [new TableRow({ children: [cell('Name', true), cell('Value', true)] }), new TableRow({ children: [cell('a'), cell('1')] })],
        }),
        new Paragraph(''),
        image(square(240, 90, 40)),
        new Paragraph('The end.'),
      ],
    },
  ],
});

writeFileSync(new URL('../fixtures/sample.docx', import.meta.url), await Packer.toBuffer(doc));
console.log('wrote fixtures/sample.docx');
