# muban-docx

Modern DOCX template engine for Java — SpEL expressions, conditional blocks,
table row replication, dynamic image replacement, and multi-format export (DOCX, PDF, HTML)
via docx4j + Apache FOP.

## Features

| Feature | Description |
| --- | --- |
| **SpEL expressions** | `${name}`, `${price * qty}`, `${price > 1000 ? 'Premium' : 'Standard'}` — powered by Spring Expression Language with a sandboxed evaluator |
| **Conditional blocks** | `#{if expr}` / `#{else}` / `#{fi}` — paragraph-level conditionals with boolean coercion |
| **Table row replication** | Automatically clone template rows for array data (`${items.name}`) |
| **Dynamic images** | Replace placeholder images via data-URI, URL, relative path, or SpEL expression in alt-text |
| **Table style flattening** | Resolves Word table conditional formatting (banding, header/footer rows, first/last columns) for PDF and HTML export fidelity |
| **PDF export** | DOCX → PDF via docx4j FO pipeline + Apache FOP, with optional password encryption and permission control |
| **HTML export** | DOCX → HTML as a self-contained ZIP micro-site (`index.html` + `index.html_files/` assets) via docx4j XSLT pipeline |
| **Locale-aware formatting** | Number and date formatting respects `Locale` (`1 234,56` for pl-PL) |
| **Zero runtime dependencies** | beyond docx4j, Spring Expression Language, and SLF4J |

## Requirements

- Java 17+
- Maven 3.8+

## Quick start

### Maven dependency

```xml
<dependency>
    <groupId>me.muban</groupId>
    <artifactId>muban-docx</artifactId>
    <version>1.0.0</version>
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

### Conditional blocks

Mark conditional paragraphs with `#{if}` / `#{else}` / `#{fi}`:

```text
#{if showDiscount}
You qualify for a ${discountPercent}% discount!
#{else}
Standard pricing applies.
#{fi}
```

- Conditions use SpEL evaluation with boolean coercion
- Each marker must be a separate paragraph
- Blocks cannot be nested

### Table row replication

For a table bound to an array, use `${arrayName.field}` in template rows:

| Item | Price |
| --- | --- |
| `${items.name}` | `${items.price}` |

The engine detects the array binding and replicates the row for each element.

### Dynamic images

Insert a placeholder image in the DOCX and set its **alt-text** to `image:{key}`, where `{key}` identifies the image to replace. The key is resolved via the following cascade:

| Alt-text example | Resolution |
| --- | --- |
| `image:facsimile` | Looks up `facsimile` in the context — the value can be a data-URI (`data:image/...;base64,...`), an HTTP/HTTPS URL, a relative path in the template package, or raw `byte[]` |
| `image:logo` | If no context value is found, the key itself (`logo`) is used as a relative path within the template ZIP |
| `image:${gender == 'F' ? 'female.png' : 'male.png'}` | SpEL expression evaluated first, then the result goes through the same resolution cascade |
| `image:assets/${department}/stamp.png` | Mixed literal + SpEL — expressions are evaluated and the resulting path is resolved |

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

## Architecture

```text
MubanDocxEngine (facade)
 ├── DocxConditionalProcessor   #{if}/{else}/{fi} block processing
 ├── DocxTableProcessor         table row replication + placeholder substitution
 ├── DocxExpressionEvaluator    SpEL sandboxed evaluation
 ├── DocxImageReplacer          dynamic image replacement
 ├── DocxTableStyleResolver     table style flattening for FO export
 ├── DocxContextBuilder         parameter + data merging
 ├── DocxExporter               DOCX save + PDF FO pipeline + HTML XSLT pipeline
 ├── DocxXmlUtils               low-level XML helpers
 └── LocaleUtils                locale parsing + number/date formatting
```

## Building

```bash
cd muban-docx
mvn clean verify
```

## License

This project is licensed under the [GNU Affero General Public License v3.0](https://www.gnu.org/licenses/agpl-3.0.html).

Copyright (c) 2025–2026 [muban.me](https://muban.me)
