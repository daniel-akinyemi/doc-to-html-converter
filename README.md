# Folio — Document → HTML

A browser-based converter that turns Microsoft Word and Google Docs content
into clean, semantic HTML. Conversion happens entirely on the client — files
and clipboard contents never leave the user's machine.

## Two ways to convert

1. **Drop a `.docx` file** — parsed in-browser with [mammoth.js].
2. **Paste content** — click *Paste content* (uses the Clipboard API) or just
   press <kbd>⌘V</kbd> / <kbd>Ctrl+V</kbd> anywhere on the page. Word and
   Google Docs put their formatted HTML on the clipboard; the app strips the
   noise and keeps the structure.

The output is shown as a live **Preview** and as raw **HTML** source, with a
one-click copy. File uploads default to the Preview tab; pastes default to
HTML.

## What's preserved

| Feature                         | Pasted content | `.docx` file |
| ------------------------------- | :------------: | :----------: |
| Headings (`<h1>`–`<h6>`)        |       ✓        |      ✓       |
| Paragraphs, lists, links        |       ✓        |      ✓       |
| Tables                          |       ✓        |      ✓       |
| Images                          |       ✓        |   ✓ (base64) |
| Bold, italic                    |       ✓        |      ✓       |
| Underline, strikethrough        |       ✓        |      ✓       |
| Superscript, subscript          |       ✓        |      ✓       |
| Inline `<code>`, `<blockquote>` |       ✓        |      ✓       |
| Text alignment                  |       ✓        |      —       |
| Text colour, highlight          |       ✓        |      —       |

For pasted content, the sanitizer reads inline CSS (`font-weight`,
`font-style`, `text-decoration`, `vertical-align`, `text-align`, `color`,
`background-color`) and either converts it to a semantic tag or preserves it
on the element. Default values (black text, white/transparent backgrounds)
are stripped to keep the output clean.

For `.docx` files, [mammoth.js] handles structural conversion. A custom
[style map][styleMap] in [`app/components/Converter.tsx`](./app/components/Converter.tsx)
opts into underline, strikethrough, code runs, and quote/caption paragraphs
that mammoth ignores by default. Paragraph alignment and font colour are
**not** preserved on this path — this is a mammoth limitation. Pasting from
the same document captures both.

## What's stripped

- MS Office namespace tags (`<o:p>`, `<w:*>`, `<v:*>`, `<m:*>`)
- `<style>`, `<script>`, `<meta>`, `<link>`, conditional comments
- Class and id attributes (after `MsoHeading*` / `MsoTitle` are promoted to
  real headings)
- Font families, font sizes, custom margins, `mso-*` properties
- Google Docs' `<b style="font-weight:normal">` document wrapper

## Running locally

```bash
pnpm install
pnpm dev
```

The dev server runs at <http://localhost:3000>.

```bash
pnpm build  # production build
pnpm start  # serve the production build
pnpm lint   # eslint
```

## Project layout

```
app/
  layout.tsx          Root layout — fonts, metadata
  page.tsx            Landing page — hero, converter section, footer
  globals.css         Tailwind 4 entry + design tokens + .prose styles
  components/
    Converter.tsx     Drop zone, paste handler, sanitizer, view toggle
public/
  mammoth.browser.min.js   Pre-built UMD bundle, lazy-loaded on first use
```

## How the conversion works

### File path

`Converter.tsx` lazy-injects `/mammoth.browser.min.js` into the page on the
first upload (cached via a module-scoped promise so it loads exactly once).
The file is read as an `ArrayBuffer` and passed to
`mammoth.convertToHtml(input, { styleMap })`. Mammoth returns clean semantic
HTML plus an array of conversion notes, which are surfaced under a
disclosure in the result header.

The pre-built UMD bundle is used instead of `import("mammoth")` because
mammoth's Node entry references `fs` and `path`, which don't survive
browser bundling.

### Paste path

When the user clicks *Paste content*, `navigator.clipboard.read()` returns
`ClipboardItem`s; the app prefers `text/html` and falls back to
`text/plain`. A document-level `paste` listener (active only when the
converter is idle, and only when the paste target isn't an `<input>` /
`<textarea>` / `contenteditable`) provides the same flow for <kbd>⌘V</kbd>.

`sanitizeWordHtml` runs the HTML through these passes:

1. Promote `MsoHeading1-6` / `MsoTitle` paragraphs to real `<h1>`–`<h6>`
   (style attributes are carried through so alignment survives).
2. Unwrap Google Docs' outer `<b style="font-weight:normal">`.
3. Drop non-content elements (`<style>`, `<script>`, etc.) and Office
   namespace tags.
4. Convert legacy presentational tags (`<b>`, `<i>`, `<font>`) to their
   semantic equivalents, keeping any inline style for the next pass.
5. For every element: read inline styles, derive semantic wrappers
   (`<strong>` / `<em>` / `<u>` / `<s>` / `<sup>` / `<sub>`), preserve a
   whitelist of styles (`text-align` on block elements, plus `color` and
   `background-color` when not default), then strip everything else.
6. Unwrap `<span>` / `<div>` left with no attributes; remove empty `<p>`s.

`plainTextToHtml` is the fallback for clipboards with only `text/plain`:
it splits on blank lines into `<p>` blocks with `<br/>` for single newlines.

## Tech

- [Next.js 16](https://nextjs.org) (App Router)
- React 19
- Tailwind CSS 4 (CSS-first config in [`app/globals.css`](./app/globals.css))
- [mammoth.js][mammoth.js] for `.docx` parsing
- Geist Sans + Geist Mono + Instrument Serif

[mammoth.js]: https://github.com/mwilliamson/mammoth.js
[styleMap]: https://github.com/mwilliamson/mammoth.js#writing-style-maps
