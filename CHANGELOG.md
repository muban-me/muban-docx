# Changelog

All notable changes to `muban-docx` are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.8.0] — 2026-11-23

### Added

- **Formatting helper `#fmt`.** The engine injects a locale-bound `DocxFormatHelper` into the SpEL
  context as the `#fmt` variable, so presentation patterns can live in the template:
  `${#fmt.date(document_date, 'd MMMM yyyy')}`, `${#fmt.number(amount, '#,##0.00')}`. The input may
  be a `TemporalAccessor`, a `java.util.Date`, or a string (ISO-8601 plus a few legacy shapes); a
  value that cannot be parsed is printed unchanged and logged. The helper exists because the
  sandbox forbids `T(...)` and constructors, so a template could not build a `DateTimeFormatter`.
- **Template variables `#{let name = expression}`.** Declarations in their own paragraphs, resolved
  sequentially in document order before conditional blocks, then removed from the output. Declared
  names live in the **SpEL variable namespace**, so they are referenced as `${#name}` in text and
  `#{if #name ...}` in conditions; the unprefixed namespace stays reserved for input data. Values
  keep their raw type (a number stays a number). `root` and `this` are reserved by SpEL and `fmt`
  by the engine — declarations using those names are ignored with a warning.
- **Unknown-variable pre-flight.** A `${#name}` (or a condition) that references a variable which
  is neither declared nor provided by the engine now leaves the placeholder visible and logs a
  warning. Without it, SpEL would resolve a missing variable to `null` and the template would
  silently print an empty string — unacceptable for legal correspondence.
- **`MubanDocxEngine.extractDeclaredVariables(WordprocessingMLPackage)`.** Lists declarations from
  the body and headers/footers without rendering, for template validation and for deriving which
  input fields a template depends on.
- Overloads carrying SpEL variables through the processing steps:
  `DocxExpressionEvaluator.evaluate(…, Map variables)`, `DocxConditionalProcessor.processConditionals(…)`,
  `DocxTableProcessor.processTable(…)`, `DocxImageReplacer.replaceImages(…)`. The previous
  signatures are kept and delegate with an empty variable map.

### Changed

- **Array-bound table rows are detected by the `<arrayKey>.<field>` reference instead of the literal
  `${arrayKey.field}` placeholder.** Previously a row whose placeholder was wrapped in an
  expression — for example `${#fmt.date(items.due_date, 'dd.MM.yyyy')}` — was not recognised as
  belonging to the array, was handled as a scalar row, and the placeholder stayed unresolved.
- `MubanDocxEngine` now resolves declarations into the SpEL variable namespace instead of writing
  them into the document context, so `${name}` and `${#name}` can never collide.
- `DocxExpressionEvaluator.extractVariableReferences` no longer reports `#`-prefixed identifiers:
  `#fmt` and `#salutation` are template internals, not input fields, so template discovery stays a
  pure view of the contract. The helper's arguments are still reported
  (`#fmt.date(document_date, ...)` → `document_date`).
- `LocaleUtils.parseLocale` builds locales with `Locale.of(...)` instead of the deprecated `Locale`
  constructors. Behaviour is unchanged — including the rejection of a fourth segment — and the
  module now compiles without deprecation warnings on Java 21.

### Documentation

- README: new sections *Formatting in templates (`#fmt`)* and *Template variables (`#{let}`)*, plus
  a note that `#fmt` is the preferred replacement for the request-side format map.

## [1.7.0] — 2026-11-16

### Added

- **Nested conditional blocks.** `#{if}` blocks can now contain other `#{if}` blocks, to any depth.
  The matcher counts `#{if}` depth instead of taking the first `#{fi}` it sees, so an inner block is
  closed by its own `#{fi}`. Previously a nested block silently corrupted the output: the inner
  `#{fi}` terminated the outer block and the remaining markers were printed literally.
- **`#{elseif expr}` branches.** A conditional block may now have any number of intermediate
  branches (alias `#{elif expr}`). Branches are evaluated top-down and the first match wins, which
  makes the construct behave like a `switch`. `DocxConditionalProcessor.ELSEIF_PATTERN` is exposed
  for tooling.

### Changed

- `DocxConditionalProcessor.processConditionals` resolves a block as a *list of branches* rather
  than an `#{if}`/`#{else}`/`#{fi}` triple. A block where no branch matches **and** there is no
  `#{else}` is now removed entirely; previously the `#{if}` and `#{fi}` markers were consumed as
  well, so the observable behaviour is unchanged.
- `findClosingMarkers(List, int)` (package-private) is now depth-aware: it returns the `#{fi}` that
  closes the scanned block, not the first one in the list.
- `extractConditionalVariables` also collects variables referenced in `#{elseif}` expressions, so
  template uploads register the right parameters.
- An `#{elseif}` placed after `#{else}` is ignored with a warning, and its body stays in the else
  branch. Duplicate `#{else}` markers are dropped the same way.
- Run merging (markers split across Word runs) moved into a new package-private helper,
  `DocxRunUtils`. The logic is unchanged and still marker-based; the processor now only deals with
  block structure and branch selection.

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
