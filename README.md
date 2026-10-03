# Docx to X Articles

A Chrome and Firefox extension. Drag a Word `.docx` onto an X Articles draft (`https://x.com/compose/articles/edit/…`) and it fills in:

- **Title**: from the paragraph styled *Title*, else a leading *Heading 1*, else the document's File → Properties title, else the file name.
- **Body**: Heading 1 → Heading; Heading 2–6 and Subtitle → Subheading; bold, italic, strikethrough, links, bulleted/numbered lists (nested), and quotes (styles *Quote*, *Intense Quote*, *Block Text*).
- **Images**: uploaded through X's own uploader and placed where they were in the document.

What changes on the way in: X Articles has no tables, underline or code blocks. Tables become `cell | cell` lines, underline is dropped, and code is kept as plain text. EMF/WMF images can't be decoded by browsers, so they're skipped. A summary card lists anything that was changed.

## How it works

| File | World | Role |
| --- | --- | --- |
| `entrypoints/x-articles.content.ts` | isolated | Drag & drop, overlay UI, parses the .docx |
| `lib/convert.ts`, `lib/blocks.ts` | isolated | mammoth → HTML → flat Draft.js-style blocks |
| `lib/images.ts` | isolated | Re-encodes non-PNG/JPEG/GIF/WebP images to PNG |
| `entrypoints/editor-bridge.content.ts`, `lib/draft.ts` | main | Writes blocks into X's Draft.js `EditorState` and uploads each image at its placeholder through X's `onFilesAdded` handler |

The two halves talk over `window.postMessage`. The main-world part reaches X's editor through React internals, an approach adapted from [xPoster](https://github.com/nevertoday/xposter) (MIT). **If X changes its editor, `lib/draft.ts` is the file to fix**: selectors are at the top, and the fiber lookups are `findDraftNode` / `findOnFilesAdded`.

## Develop

```sh
npm install
npm run dev            # Chrome with the extension loaded
npm run dev:firefox    # Firefox (128+)
npm test               # unit tests: docx → blocks (fixtures/sample.docx)
npm run test:e2e       # headless Chrome: full import into a Draft.js stand-in for X's editor (harness/)
npm run compile        # type-check
```

`npm run fixtures` regenerates `fixtures/sample.docx`. `npm run icons` redraws `public/icon/*`.

## Publish

```sh
npm run zip            # .output/*-chrome.zip  → Chrome Web Store
npm run zip:firefox    # .output/*-firefox.zip + *-sources.zip → addons.mozilla.org
```

- **Firefox:** before the first AMO upload, change `browser_specific_settings.gecko.id` in `wxt.config.ts` to an ID you own. Upload the sources zip when AMO asks for source code, since the bundle is minified. `web-ext lint` warns about `Function(...)` inside mammoth's bundled dependencies (underscore, a setImmediate polyfill). That code is third-party, and the extension never evaluates strings.
- **Both stores:** use `PRIVACY.md` as the privacy policy. The only permission is the content-script match on `x.com` / `twitter.com`.
