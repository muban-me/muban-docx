# muban-docx

Modern DOCX template engine for Java — SpEL expressions, conditional blocks,
table row replication, dynamic image replacement, and PDF export via docx4j + Apache FOP.

## Features

| Feature | Description |
| --- | --- |
| **SpEL expressions** | `${name}`, `${price * qty}`, `${amount.format('%.2f')}` — powered by Spring Expression Language with a sandboxed evaluator |
| **Conditional blocks** | `#{if expr}` / `#{else}` / `#{fi}` — paragraph-level conditionals with boolean coercion |
| **Table row replication** | Automatically clone template rows for array data (`${items.name}`) |
| **Dynamic images** | Replace placeholder images via data-URI, URL, relative path, or SpEL expression in alt-text |
| **Table style flattening** | Resolves Word table conditional formatting (banding, header/footer rows, first/last columns) for FO export fidelity |
| **PDF export** | DOCX → PDF via docx4j FO pipeline + Apache FOP, with optional password encryption and permission control |
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

Insert a placeholder image in the DOCX and set its **alt-text** to one of:

| Alt-text format | Description |
| --- | --- |
| `${fieldName}` | Resolved from context (base64, data-URI, or relative path) |
| `url:https://...` | Fetched from URL at generation time |
| `path:images/logo.png` | Relative to the asset directory |
| `spel:condition ? 'a.png' : 'b.png'` | SpEL expression returning a path or data-URI |

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

## Architecture

```text
MubanDocxEngine (facade)
 ├── DocxConditionalProcessor   #{if}/{else}/{fi} block processing
 ├── DocxTableProcessor         table row replication + placeholder substitution
 ├── DocxExpressionEvaluator    SpEL sandboxed evaluation
 ├── DocxImageReplacer          dynamic image replacement
 ├── DocxTableStyleResolver     table style flattening for FO export
 ├── DocxContextBuilder         parameter + data merging
 ├── DocxExporter               DOCX save + PDF FO pipeline
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
