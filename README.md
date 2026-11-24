# muban-docx

Modern DOCX template engine for Java — SpEL expressions, conditional blocks,
table row replication, dynamic image replacement, and multi-format export (DOCX, PDF, HTML, TXT, PNG)
via docx4j + Apache FOP.

## Features

| Feature | Description |
| --- | --- |
| **SpEL expressions** | `${name}`, `${price * qty}`, `${price > 1000 ? 'Premium' : 'Standard'}` — powered by Spring Expression Language with a sandboxed evaluator |
| **Template variables** | `#{let name = expr}` declarations, referenced as `${#name}` and in `#{if}` conditions — the DOCX counterpart of JasperReports report variables |
| **Formatting helper** | `#fmt.date` and `#fmt.number` — date and number patterns live in the template instead of the input data |
| **Rich text values** | `${html:footer}` renders a whitelisted HTML subset (`<br>`, `<b>`, `<i>`, `<u>`, `<sup>`, `<sub>`, entities) as Word runs — for HTML fragments coming from a CMS or rules engine |
| **Conditional blocks** | `#{if expr}` / `#{elseif expr}` / `#{else}` / `#{fi}` — paragraph-level branches with boolean coercion, nesting supported |
| **Table row replication** | Automatically clone template rows for array data (`${items.name}`) |
| **Dynamic images** | Replace placeholder images via data-URI, URL, relative path, SpEL expression, or generated **QR code** in alt-text |
| **Table style flattening** | Resolves Word table conditional formatting (banding, header/footer rows, first/last columns) for PDF and HTML export fidelity |
| **PDF export** | DOCX → PDF via docx4j FO pipeline + Apache FOP, with optional password encryption and permission control |
| **PNG export** | DOCX → PNG via docx4j FO pipeline + Apache FOP, one `page_N.png` per page packaged as a ZIP archive — no intermediate PDF |
| **HTML export** | DOCX → HTML as a self-contained ZIP micro-site (`index.html` + `index.html_files/` assets) via docx4j XSLT pipeline |
| **TXT export** | Plain-text extraction with configurable line separator, trailing-whitespace trimming, and soft word-wrap |
| **Format map** | Per-placeholder format patterns (`#,##0.00`, `dd.MM.yyyy`) applied after SpEL evaluation via `.formatMap()` |
| **Soft line break handling** | Runs containing `w:br` (soft line breaks) are split before processing, preventing placeholder corruption |
| **Locale-aware formatting** | Number and date formatting respects `Locale` (`1 234,56` for pl-PL) |
| **Zero runtime dependencies** | beyond docx4j, Spring Expression Language, SLF4J and ZXing (QR codes only) |

## Requirements

- Java 21+
- Maven 3.8+
- PDF and PNG export require the optional `docx4j-export-fo` dependency on the classpath

## Quick start

### Maven dependency

```xml
<dependency>
    <groupId>me.muban</groupId>
    <artifactId>muban-docx</artifactId>
    <version>1.7.0</version>
</dependency>
```

### Generate a document

```java
import me.muban.docx.MubanDocxEngine;

Map<String, Object> data = Map.of(
    "recipientName", "Jan Kowalski",
    "amount", 1500.50,
    "items", List.of(
        Map.of("name", "Widget A", "price", 29.99),
        Map.of("name", "Widget B", "price", 14.99)
    )
);

String outputPath = MubanDocxEngine.builder()
    .template(new File("invoice.docx"))
    .data(data)
    .locale(Locale.forLanguageTag("pl-PL"))
    .outputDir("/tmp/output/")
    .outputFormat("pdf")
    .build()
    .generate();
```

### Format map (post-evaluation formatting)

Apply format patterns to specific placeholders after SpEL evaluation:

```java
Map<String, String> formats = Map.of(
    "amount", "#,##0.00",
    "invoiceDate", "dd.MM.yyyy"
);

String outputPath = MubanDocxEngine.builder()
    .template(new File("invoice.docx"))
    .data(data)
    .formatMap(formats)
    .locale(Locale.forLanguageTag("pl-PL"))
    .outputDir("/tmp/output/")
    .outputFormat("pdf")
    .build()
    .generate();
// amount renders as "1 500,50", invoiceDate as "15.01.2025"
```

Format patterns follow `DecimalFormat` syntax for numbers and `DateTimeFormatter`
syntax for temporal values. They are applied only to simple key placeholders
(e.g. `${amount}`), not to SpEL expressions.

> For new templates prefer `#fmt` (see *Formatting in templates*) — the pattern then lives in the
> document, applies inside expressions as well, and needs no wiring in the request.

### In-memory processing (no export)

```java
WordprocessingMLPackage result = MubanDocxEngine.builder()
    .template(inputStream)
    .data(data)
    .build()
    .process();

// Save or further manipulate the package yourself
result.save(new File("result.docx"));
```

> **Note:** `process()` does **not** call `VariablePrepare.prepare()` on the
> package. For DOCX files loaded from disk (which may contain split runs, rsid
> attributes, and bookmarks), call `VariablePrepare.prepare(wordPackage)` before
> invoking `process()`. The `generate()` method calls `prepare()` automatically.

## Template syntax

### Placeholders

Use `${expression}` anywhere in the document body, headers, or footers:

```text
Dear ${recipientName},

Your order total is ${amount} PLN.
```

SpEL expressions are fully supported:

```text
${price > 1000 ? 'Premium' : 'Standard'}
${name.toUpperCase()}
${items.size()}
```

### Formatting in templates (`#fmt`)

Date and number patterns belong to the template, not to the data. The engine injects a
locale-bound helper into the SpEL context as `#fmt`, so the pattern is written where the value
is rendered:

```text
Payment due: ${#fmt.date(due_date, 'd MMMM yyyy')}.
Amount: ${#fmt.number(outstanding_amount, '#,##0.00')}.
```

| Call | Result (en-US) |
| --- | --- |
| `${#fmt.date(d, 'yyyy-MM-dd')}` | `2026-11-23` |
| `${#fmt.date(d, 'd MMMM yyyy')}` | `23 November 2026` |
| `${#fmt.date(d, 'LLLL yyyy')}` | `November 2026` — stand-alone form |
| `${#fmt.number(n, '#,##0.00')}` | `1,234.50` |

The locale passed to the engine decides month names and separators. `MMMM` follows the
format-wide form and `LLLL` the stand-alone form; the difference matters for languages that inflect
month names (Slavic ones, for instance).

- The input may be a `LocalDate` / `LocalDateTime`, a `java.util.Date`, or a string. ISO-8601 is
  the expected string form; `dd.MM.yyyy`, `dd-MM-yyyy`, `dd/MM/yyyy` and `yyyy-MM-dd HH:mm[:ss]`
  are accepted for legacy data.
- `#fmt.date(value)` and `#fmt.number(value)` default to `dd.MM.yyyy` and `#,##0.00`.
- A value that cannot be parsed is printed unchanged and logged, so the problem is visible in the
  document instead of turning into a blank cell.
- Works anywhere an expression works — including branch conditions and replicated table rows.

Why a helper and not a static call? The evaluator sandbox forbids `T(...)` and constructors, so a
template cannot build its own formatter:
`T(java.time.format.DateTimeFormatter).ofPattern(...)` fails with `EL1005E`.

**Scope.** `#fmt` is deliberately domain-neutral: it knows patterns and locales, not business
vocabulary. Anything that needs a dictionary — amounts in words, month names in the locative case,
declension of nouns or forms of address — stays out of the library. Wording that depends on a data
value is written as an ordinary expression, which also keeps the text visible next to its context:

```text
${customer_type == 'CORPORATE' ? 'the company has paid' : 'you have paid'}
#{let salutation = addressee_gender == 'M' ? 'Dear Mr,' : (addressee_gender == 'F' ? 'Dear Ms,' : 'Dear Sir or Madam,')}
```

This covers words whose form depends **only** on the value — verb forms, adjectives, pronouns,
forms of address. Lexical variation (a job title that has to decline, for example) is not derivable
and must come from the data.

### Template variables (`#{let}`)

A declaration computes a value once and reuses it. Declarations are paragraphs of their own, like
conditional markers, and are removed from the output:

```text
#{let salutation = addressee_gender == 'M' ? 'Dear Mr,' : (addressee_gender == 'F' ? 'Dear Ms,' : 'Dear Sir or Madam,')}
#{let due_long = #fmt.date(payment_due_date, 'd MMMM yyyy')}
#{let demand = 'the sum of ' + #fmt.number(outstanding_amount, '#,##0.00')}
```

```text
${#salutation}
Payment of ${#demand} is due by ${#due_long}.
```

Declared names live in the **SpEL variable namespace**, so they are always referenced with a `#`:
`${#salutation}` in text, `#{if #salutation == ...}` in conditions. That makes the template's
contract readable — by a person and by tooling — without knowing anything about `#{let}`:

| Syntax | Meaning |
| --- | --- |
| `${case_number}` | input data supplied by the caller |
| `${#salutation}`, `${#fmt.date(...)}` | computed inside the template |
| `#{if ...}` / `#{let ...}` / `#{fi}` | directives, not values |

Rules:

- **Document order.** A declaration may use the ones before it — with `#`, so
  `#{let total = #net + #vat}`. A forward reference fails and leaves the variable undefined,
  so there are no cycles to detect.
- **Scope.** The body and headers/footers share one namespace.
- **Raw value.** The evaluated object is stored as-is — a number stays a number — so
  `#fmt.number(#total, ...)` can format it later and `${#total}` formats it at output time.
- **Separate namespace.** A declaration cannot shadow input data: `${name}` reads the data,
  `${#name}` reads the declaration, and neither ever sees the other.
- **Reserved names.** `root` and `this` belong to the SpEL runtime and `fmt` to the engine; a
  declaration using one of them is ignored with a warning.
- **Failure is loud.** An expression that cannot be evaluated, or a reference to an undeclared
  `#variable`, leaves `${...}` visible in the output instead of rendering an empty string.
- **Not inside replicated table rows.** Declarations are resolved before row replication, so a
  declaration placed in a template row would run once. Use an expression in the row instead.

Because declarations are resolved before conditional blocks, `#{let}` cannot be assigned
conditionally — a declaration inside an `#{if}` block would run unconditionally. Use a ternary
inside the declaration:

```text
#{let greeting = addressee_gender == 'F' ? 'Dear Ms,' : 'Dear Mr,'}
```

Declarations can be listed without rendering, which is useful for validating a template or for
diagnosing why a `${#name}` stayed unresolved:

```java
Map<String, String> declared = MubanDocxEngine.extractDeclaredVariables(wordPackage);
// {salutation=addressee_gender == 'M' ? ... , due_long=#fmt.date(payment_due_date, 'd MMMM yyyy'), ...}
```

### Rich text values

Values managed outside Word (legal footers, clauses, marketing texts) are often stored
as HTML. Prefix the placeholder with `html:` to render a whitelisted HTML subset
instead of printing the markup literally:

```text
${html:marketing_footer}
```

| Input | Result |
| --- | --- |
| `<br>`, `<br/>` | line break |
| `<b>`, `<strong>` | bold |
| `<i>`, `<em>` | italic |
| `<u>` | single underline |
| `<sup>`, `<sub>` | superscript / subscript |
| `<p>`, `<div>`, `<li>`, `<tr>` | paragraph boundary (single line break) |
| `&amp;`, `&lt;`, `&gt;`, `&nbsp;`, numeric entities | decoded characters |
| any other tag | stripped, inner text kept |

Rules:

- **Opt-in only.** Without the `html:` prefix the value is inserted as plain text —
  literal `<b>` stays `<b>`.
- **Formatting is additive.** The template's run formatting is preserved; HTML only
  toggles bold/italic/underline/vertical alignment on top of it. Use `&lt;b&gt;` inside
  a rich value to show literal markup.
- **Bold always renders, from a bold file or synthesised.** `<b>` sets `w:b`. If the family
  provides a bold face the renderer uses it; if it does not ("… Light" families) the regular
  face's outline is thickened instead, at identical glyph advances — so bold never reflows.
- **CSS, links, images and tables are intentionally not supported.** Presentation must
  come from the template, not from the data; images have their own channel (`image:` alt-text).
- **Never fails.** Malformed input is treated as text, unknown tags are stripped, and values
  longer than 20 000 characters degrade to plain text.

The prefix works anywhere a placeholder works, including headers and footers, and can
wrap any expression: `${html:footer != null ? footer : ''}`.

If a template cannot be edited, the same behaviour can be forced from code with
`.richTextKeys(Set.of("marketing_footer"))` — the in-template marker is the recommended
mechanism, because it keeps the intent versioned with the document.

### Conditional blocks

Mark conditional paragraphs with `#{if}` / `#{elseif}` / `#{else}` / `#{fi}`:

```text
#{if showDiscount}
You qualify for a ${discountPercent}% discount!
#{else}
Standard pricing applies.
#{fi}
```

`#{elseif expr}` (alias `#{elif expr}`) adds intermediate branches; they are tested top-down and
the first match wins. `#{else}` must be the last branch.

```text
#{if type == 'A'}
Alpha
#{elseif type == 'B'}
Beta
#{else}
Other
#{fi}
```

- Conditions use SpEL evaluation with boolean coercion
- Each marker must be a separate paragraph
- Blocks may be nested (the matcher counts `#{if}` depth); keep nesting shallow for readability
- A whole block with no matching branch and no `#{else}` is removed

### Table row replication

For a table bound to an array, use `${arrayName.field}` in template rows:

| Item | Price |
| --- | --- |
| `${items.name}` | `${items.price}` |

The engine detects the array binding and replicates the row for each element.

A row is recognised by any reference to `<arrayName>.<field>`, so placeholders may be wrapped in
expressions — `${#fmt.date(items.due_date, 'dd.MM.yyyy')}` still replicates the row and formats each
clone with its own value.

### Dynamic images

Insert a placeholder image in the DOCX and set its **alt-text** to `image:{key}`, where `{key}`
identifies the image to replace. The key is resolved via the following cascade:

| Alt-text example | Resolution |
| --- | --- |
| `image:facsimile` | Looks up `facsimile` in the context — the value can be a data-URI (`data:image/...;base64,...`), an HTTP/HTTPS URL, a relative path in the template package, or raw `byte[]` |
| `image:logo` | If no context value is found, the key itself (`logo`) is used as a relative path within the template ZIP |
| `image:${gender == 'F' ? 'female.png' : 'male.png'}` | SpEL expression evaluated first, then the result goes through the same resolution cascade |
| `image:assets/${department}/stamp.png` | Mixed literal + SpEL — expressions are evaluated and the resulting path is resolved |

#### QR codes

A context value prefixed with `qr:` is rendered as a PNG QR code in place of the placeholder —
no need to generate the image upstream:

```java
Map<String, Object> data = Map.of(
    "qr_code", "qr:https://www.example.com/pay?nr=123456&amount=100.00"
);
```

```text
Alt-text: image:qr_code
```

Codes use UTF-8 and are rendered oversampled (300 px minimum) and scaled down by the picture
frame in the template, so the placeholder image **must be square** — otherwise the code will be
distorted. Because the image part keeps its original content type, the placeholder must be a
**PNG**.

Rendering can be tuned per image with a query string appended to the alt-text key:

```text
Alt-text: image:qr_code?ec=H&margin=4&scale=8&bg=transparent
```

| Option | Values | Default | Meaning |
| --- | --- | --- | --- |
| `ec` | `L`, `M`, `Q`, `H` | `M` | Error correction level (~7 / 15 / 25 / 30 % recovery) |
| `margin` | `0`–`16` | `2` | Quiet zone, in modules |
| `scale` | `1`–`40` | auto | Pixels per module (disables auto-scaling) |
| `minPx` | `60`–`4000` | `300` | Minimum edge length when scaling automatically |
| `bg` | `white`, `transparent`, `#RRGGBB` | `white` | Background colour |

Unknown options and out-of-range values are ignored with a warning, and a query that is not a
well-formed `name=value` list is treated as part of the key — a typo in a template cannot break
document generation.

**Background.** Opaque white is the default on purpose: a QR code needs a light quiet zone
around the symbol, and a transparent background only stays light if nothing coloured sits behind
the picture. Use `bg=transparent` only when the template guarantees a white backdrop — note that
even then the transparent pixels keep white RGB under the alpha channel, so pipelines that drop
the alpha channel still see a light background rather than black.

Blank payloads, and any encoding failure, leave the original template image untouched.
The encoder is available directly as `QrCodeEncoder.toPng("text")` or
`QrCodeEncoder.toPng("text", QrCodeEncoder.QrOptions.parse("ec=H&bg=transparent"))`.

## PDF security

```java
PdfExportOptions security = PdfExportOptions.builder()
    .ownerPassword("secret")
    .userPassword("reader")
    .canPrint(true)
    .canCopy(false)
    .encryptionKeyLength(256)
    .build();

MubanDocxEngine.builder()
    .template(new File("template.docx"))
    .data(data)
    .outputFormat("pdf")
    .pdfOptions(security)
    .pdfSecurityCallback((path, opts) -> {
        // Apply PDF encryption using your preferred library (e.g. PDFBox)
    })
    .build()
    .generate();
```

## HTML export

HTML export produces a ZIP archive containing `index.html` and an
`index.html_files/` directory with embedded images. The structure is
compatible with the S2S convention used by JasperReports HTML output.

```java
String zipPath = MubanDocxEngine.builder()
    .template(new File("template.docx"))
    .data(data)
    .locale(Locale.forLanguageTag("pl-PL"))
    .outputDir("/tmp/output/")
    .outputFormat("html")
    .build()
    .generate();
// zipPath → /tmp/output/<uuid>.zip
// Contents:
//   <uuid>/index.html
//   <uuid>/index.html_files/image1.png
```

### Image modes

`DocxExporter.exportHtml(pkg, outputDir, HtmlExportOptions)` selects how images
are represented in the HTML:

| Mode | HTML `src` | ZIP contents | Use case |
| ---- | ---------- | ------------ | -------- |
| `FILES` (default) | `index.html_files/image1.png` | image files | web hosting, self-contained micro-site |
| `DATA_URI` | `data:image/png;base64,...` | no image files | single-file HTML; poorly supported by email clients |
| `CID` | `cid:<id>@muban` | image files + `images.json` | email bodies assembled as `multipart/related` |
| `NONE` | empty | no image files | HTML-only output; the caller strips the `<img>` elements |

In `CID` mode `images.json` maps every Content-ID to its file, so the email
sender never has to parse the HTML:

```json
{
  "version": 1,
  "images": [
    {
      "contentId": "8f14e45f...1.image1.png@muban",
      "path": "index.html_files/image1.png",
      "contentType": "image/png",
      "size": 1204
    }
  ]
}
```

Paths in the manifest are relative to `index.html`. `contentId` is written
without angle brackets — use `Content-ID: <contentId>` on the matching MIME
part.

## PNG export

PNG export renders the document's XSL-FO layout directly with Apache FOP — no
intermediate PDF is created. Each page is written as `page_N.png` and packaged
into a ZIP archive.

```java
String zipPath = MubanDocxEngine.builder()
    .template(new File("template.docx"))
    .data(data)
    .outputDir("/tmp/output/")
    .outputFormat("png")
    .pngOptions(new PngExportOptions(2.0f))  // 2× resolution (144 DPI)
    .build()
    .generate();
// zipPath → /tmp/output/<uuid>.zip
// Contents:
//   <uuid>/page_1.png
//   <uuid>/page_2.png
//   ...
```

### PNG export options

| Option | Default | Description |
| --- | --- | --- |
| `zoomRatio` | `1.0` | Output resolution multiplier. `1.0` = native page size (~72 DPI), `2.0` = 2× (~144 DPI), `4.0` = 4× (~288 DPI). Must be a positive, finite number. |

Higher zoom ratios produce sharper images but increase memory usage and file
size quadratically (2× zoom = 4× pixel count). Pages are rendered sequentially.

## TXT export

Plain-text export walks all `w:p` paragraphs in the DOCX body (including tables)
and writes each paragraph's text separated by a configurable line separator.
Output is always UTF-8.

```java
// Basic — system line separator, no trimming
String txtPath = MubanDocxEngine.builder()
    .template(new File("template.docx"))
    .data(data)
    .outputDir("/tmp/output/")
    .outputFormat("txt")
    .build()
    .generate();
```

### TXT export options

| Option | Default | Description |
| --- | --- | --- |
| `lineSeparator` | system default | String inserted between paragraphs (`"\n"`, `"\r\n"`, `" \| "`, …) |
| `trimLineRight` | `false` | Strip trailing whitespace from each paragraph (removes Word's trailing spaces) |
| `pageWidthInChars` | none (no wrap) | Soft word-wrap width — lines are broken at spaces so no line exceeds this many characters. Words longer than the limit are kept intact. |

```java
TxtExportOptions opts = new TxtExportOptions("\n", true, 80);
String txtPath = MubanDocxEngine.builder()
    .template(new File("template.docx"))
    .data(data)
    .outputDir("/tmp/output/")
    .outputFormat("txt")
    .txtOptions(opts)
    .build()
    .generate();
```

## Architecture

```text
MubanDocxEngine (facade)
 ├── DocxConditionalProcessor   #{if}/#{elseif}/#{else}/#{fi} blocks, nesting supported
 ├── DocxTableProcessor         table row replication + placeholder substitution
 ├── DocxRichText               whitelisted HTML → Word runs for ${html:...}
 ├── DocxExpressionEvaluator    SpEL sandboxed evaluation
 ├── DocxImageReplacer          dynamic image replacement
 ├── QrCodeEncoder              QR code rendering for qr: values
 ├── DocxTableStyleResolver     table style flattening for FO export
 ├── DocxContextBuilder         parameter + data merging
 ├── DocxExporter               DOCX save + PDF FO pipeline + PNG FOP pipeline + HTML XSLT pipeline + TXT extraction
 ├── PngFoRenderer              page rasterisation for PNG export
 ├── DocxXmlUtils               low-level XML helpers
 ├── DocxRunUtils               run re-joining (markers split by Word) + formatting comparison
 └── LocaleUtils                locale parsing + number/date formatting
```

## Building

```bash
cd muban-docx
mvn clean verify
```

## License

muban-docx is **dual-licensed**.

### Open source — [GNU Affero General Public License v3.0](https://www.gnu.org/licenses/agpl-3.0.html)

You may use, modify and distribute this library under AGPL-3.0, including in commercial
projects. If you do, the combined work must be made available under AGPL-3.0 — this also
applies when the software is offered to users over a network (AGPL-3.0, section 13).

### Commercial license

If you want to embed muban-docx in a proprietary product, without the obligations of
AGPL-3.0, a separate commercial license is available from the copyright holder.
Contact: [muban.me](https://muban.me).

### Contributions

By submitting a contribution to this project you agree that it may be distributed under
the licenses above, and you grant the copyright holder a perpetual, irrevocable,
worldwide, royalty-free license to use, modify, sublicense and relicense your
contribution as part of this project.

The full text of AGPL-3.0 is included in the [LICENSE](LICENSE) file.

Copyright (c) 2025–2026 [muban.me](https://muban.me)
