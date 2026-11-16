# Changelog

All notable changes to `muban-docx` are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.6.0] — 2026-11-16

### Added

- **Rich text placeholders.** A placeholder prefixed with `html:` (for example `${html:marketing_footer}`)
  renders a whitelisted HTML subset as Word runs instead of printing the markup literally:
  `<br>`/`<br/>` (line break), `<b>`/`<strong>`, `<i>`/`<em>`, `<u>`, `<sup>`/`<sub>`,
  `<p>`/`<div>`/`<li>`/`<tr>` (paragraph boundary), named and numeric entities.
  Unknown tags are stripped, malformed input degrades to text, and values longer than
  20 000 characters fall back to plain text. New class: `DocxRichText`.
- `MubanDocxEngine.RICH_TEXT_PREFIX` — public constant for the `html:` marker.
- `MubanDocxEngine.Builder.richTextKeys(Set<String>)` — forces rich rendering for named keys,
  for templates that cannot be edited.
- `DocxXmlUtils.buildRun(RPr, List<Object>)` and `DocxXmlUtils.textRun(RPr, String)` are now public.

### Changed

- Placeholder substitution is split into two paths: the original single-run substitution
  (output unchanged, used for every unmarked placeholder) and a run-rebuilding path used only
  when a rich-text placeholder is present.
- Rich text is processed in the document body, headers, footers and table cells alike.

## [1.5.0] — 2026-11-09

### Added

- **PNG export.** `outputFormat("png")` renders the document's XSL-FO layout directly with Apache FOP
  — no intermediate PDF — and packages one `page_N.png` per page into a ZIP archive.
- `PngExportOptions` with a `zoomRatio` output-resolution multiplier (`1.0` = native page size,
  `2.0` = 2×, `4.0` = 4×).
- New class: `PngFoRenderer`.

### Changed

- `MubanDocxEngine.Builder` accepts `pngOptions(PngExportOptions)`.

## [1.4.0] — 2026-11-02

### Added

- **QR codes in image placeholders.** A context value prefixed with `qr:` is rendered as a PNG
  QR code in place of the placeholder image — no need to generate the image upstream.
- Per-image QR options through an alt-text query string:
  `image:qr_code?ec=H&margin=4&scale=8&minPx=300&bg=transparent`.
- Transparent QR background (`bg=transparent`), keeping white RGB under the alpha channel.
- New class: `QrCodeEncoder`.

### Changed

- Image resolution cascade extended with the `qr:` scheme (evaluated after data-URIs and before
  HTTP URLs and package-relative paths).

### Dependencies

- Added `com.google.zxing:core:3.5.3`.
