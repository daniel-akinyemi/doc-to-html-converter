# Folio — Document → HTML

A browser-based tool that turns Microsoft Word and Google Docs content into
clean, semantic HTML. Everything runs on the client — files and clipboard
contents never leave the user's machine.

Two pages:

- **`/` — Convert** — a whole document → one HTML block.
- **`/split` — Split** — a structured document → five labelled HTML sections.

## Convert (`/`)

Two ways in:

1. **Drop a `.docx` file** — parsed in-browser with [mammoth.js].
2. **Paste content** — click *Paste content* (uses the Clipboard API) or just
   press <kbd>⌘V</kbd> / <kbd>Ctrl+V</kbd> anywhere on the page. Word and
   Google Docs put their formatted HTML on the clipboard; the app strips the
   noise and keeps the structure.

The output is shown as a live **Preview** and as raw **HTML** source, with a
one-click copy. File uploads default to the Preview tab; pastes default to
HTML.

## Split (`/split`)

Takes a structured document (built for product "buyer's guide"-style docs) and
breaks it into five sections, each as its own copyable HTML block, shown in
tabs:

| Section                      | Built from these doc headings                                                                                |
| ---------------------------- | ------------------------------------------------------------------------------------------------------------ |
| **Product overview**         | the intro text + *What makes this product stand out?* + *What is this product made of?* + *What is it used for?* + *Why you'll appreciate this product* |
| **Specifications**           | *Specs and compliance you should know*                                                                       |
| **Sizing & selection**       | *How to choose the right size and type* + *Explore other options*                                            |
| **Care & storage**           | *Care, storage, and disposal*                                                                                |
| **Frequently asked questions** | *Frequently asked questions*                                                                                |

The original headings are kept as `<h3>` sub-headings inside each section. The
two "Ready to upgrade your supply?" marketing blocks are dropped. A section
with no matching content shows "Not found in this document".

**Single doc vs. bundle.** If the file holds several `… – Complete Buyer's
Guide` titles, each one is detected as a separate product (split on its own,
selectable from a picker at the top). If there's just one, you get a flat
5-tab view. Unstructured docs (no recognised headings) land entirely in
*Product overview* with the other four tabs empty.

Input methods are the same as the converter: file drop, *Paste content*, or
<kbd>⌘V</kbd>.

## What's preserved

| Feature                         | Pasted content | `.docx` file |
| ------------------------------- | :------------: | :----------: |
| Headings (`<h3>`–`<h6>`)        |       ✓        |      ✓       |
| Paragraphs, lists, links        |       ✓        |      ✓       |
| Tables                          |       ✓        |      ✓       |
| Images                          |       ✓        |   ✓ (base64) |
| Bold, italic                    |       ✓        |      ✓       |
| Underline, strikethrough        |       ✓        |      ✓       |
| Superscript, subscript          |       ✓        |      ✓       |
| Inline `<code>`, `<blockquote>` |       ✓        |      ✓       |
| Text colour, highlight          |       ✓        |      —       |

**Heading levels are clamped to `<h3>` maximum.** Anything that came in as
`<h1>` or `<h2>` (or Word's *Title* / *Heading 1* / *Heading 2* styles) is
emitted as `<h3>`. `<h3>`–`<h6>` pass through unchanged.

**Text alignment is not preserved** — output is always left-aligned.
`text-align` is dropped from inline styles, `<center>` tags are unwrapped,
and centered Word/Docs paragraphs come out left.

For pasted content, the sanitizer reads inline CSS (`font-weight`,
`font-style`, `text-decoration`, `vertical-align`, `color`,
`background-color`) and either converts it to a semantic tag or preserves it
on the element. Default values (black text, white/transparent backgrounds)
are stripped to keep the output clean.

For `.docx` files, [mammoth.js] handles structural conversion. A custom
[style map][styleMap] in [`app/lib/docx.ts`](./app/lib/docx.ts) opts into
underline, strikethrough, code runs, and quote/caption paragraphs that
mammoth ignores by default. Paragraph alignment and font colour are **not**
preserved on this path — this is a mammoth limitation. Pasting from the same
document captures both. (The split tool runs documents through the same
conversion before slicing them up, so the same rules apply.)

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
  layout.tsx           Root layout — fonts, metadata
  page.tsx             "Convert" page — hero, converter, footer
  split/page.tsx       "Split" page — hero, splitter, footer
  globals.css          Tailwind 4 entry + design tokens + .prose styles
  lib/
    docx.ts            mammoth loader, style map, paste sanitiser, splitDocument()
  components/
    Converter.tsx      Convert page UI
    Splitter.tsx       Split page UI — product picker + section tabs
    UploadZone.tsx     Shared: drop zone + paste card + ⌘V listener
    ViewToggle.tsx     Shared: Preview / HTML toggle
    SiteChrome.tsx     Shared header & footer
public/
  mammoth.browser.min.js   Pre-built UMD bundle, lazy-loaded on first use
```

## How the conversion works

### File path

`loadMammoth()` (in [`app/lib/docx.ts`](./app/lib/docx.ts)) lazy-injects
`/mammoth.browser.min.js` into the page on the first upload (cached via a
module-scoped promise so it loads exactly once). The file is read as an
`ArrayBuffer` and passed to `mammoth.convertToHtml(input, { styleMap })`.
Mammoth returns clean semantic HTML plus an array of conversion notes, which
are surfaced under a disclosure in the converter's result header.

The pre-built UMD bundle is used instead of `import("mammoth")` because
mammoth's Node entry references `fs` and `path`, which don't survive
browser bundling.

### Paste path

When the user clicks *Paste content*, `navigator.clipboard.read()` returns
`ClipboardItem`s; the app prefers `text/html` and falls back to
`text/plain`. A document-level `paste` listener (active only while the
upload zone is shown, and only when the paste target isn't an `<input>` /
`<textarea>` / `contenteditable`) provides the same flow for <kbd>⌘V</kbd>.
This lives in [`app/components/UploadZone.tsx`](./app/components/UploadZone.tsx),
shared by both pages.

`sanitizeWordHtml` (in [`app/lib/docx.ts`](./app/lib/docx.ts)) runs the HTML
through these passes:

1. Promote `MsoHeading1-6` / `MsoTitle` paragraphs to real headings (clamped
   to `<h3>`).
2. Unwrap Google Docs' outer `<b style="font-weight:normal">`.
3. Drop non-content elements (`<style>`, `<script>`, etc.) and Office
   namespace tags.
4. Convert legacy presentational tags (`<b>`, `<i>`, `<font>`, `<center>`) to
   semantic equivalents (or unwrap), keeping any inline style for the next pass.
5. For every element: read inline styles, derive semantic wrappers
   (`<strong>` / `<em>` / `<u>` / `<s>` / `<sup>` / `<sub>`), preserve `color`
   and `background-color` when not default, then strip everything else.
6. Unwrap `<span>` / `<div>` left with no attributes; remove empty `<p>`s.
7. Clamp any remaining `<h1>` / `<h2>` to `<h3>`.

`plainTextToHtml` is the fallback for clipboards with only `text/plain`:
it splits on blank lines into `<p>` blocks with `<br/>` for single newlines.

### Splitting

`splitDocument(html, fallbackTitle)` parses the converted/sanitised HTML,
scans the top-level blocks, and:

- Detects product boundaries by heading/paragraph text containing
  "Complete Buyer's Guide" — one or zero ⇒ single product, more ⇒ a bundle
  (walking back over `SKU` / `ID` metadata lines to keep each product's
  preamble with the right product).
- Within each product, walks the blocks in order and routes each into one of
  the five buckets (or `drop`) using a text-match table on recognised section
  headings; content before the first recognised heading goes to *Product
  overview*.
- Recognised section headings are re-emitted as `<h3>`; each bucket's HTML is
  serialised, with `<h1>`/`<h2>` clamped to `<h3>` and empty `<p>`s removed.

## Tech

- [Next.js 16](https://nextjs.org) (App Router)
- React 19
- Tailwind CSS 4 (CSS-first config in [`app/globals.css`](./app/globals.css))
- [mammoth.js][mammoth.js] for `.docx` parsing
- Geist Sans + Geist Mono + Instrument Serif

[mammoth.js]: https://github.com/mwilliamson/mammoth.js
[styleMap]: https://github.com/mwilliamson/mammoth.js#writing-style-maps
