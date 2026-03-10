package me.muban.docx;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * High-level facade for the muban-docx template engine.
 *
 * <p>This is the main entry point for consumers. It orchestrates the full template
 * processing pipeline:
 *
 * <ol>
 *   <li><b>Load</b> — open a DOCX file or stream</li>
 *   <li><b>Build context</b> — merge parameters and data into an evaluation context</li>
 *   <li><b>Conditional blocks</b> — evaluate {@code #{if}}/#{else}/{@code #{fi}} markers</li>
 *   <li><b>Table row replication</b> — clone template rows for data arrays</li>
 *   <li><b>Placeholder substitution</b> — replace {@code ${expr}} with values</li>
 *   <li><b>Image replacement</b> — swap placeholder images with dynamic content</li>
 *   <li><b>Export</b> — save to DOCX or PDF</li>
 * </ol>
 *
 * <p><b>Quick start:</b></p>
 * <pre>{@code
 * Map<String, Object> data = Map.of(
 *     "recipientName", "Jan Kowalski",
 *     "amount", 1500.50,
 *     "items", List.of(
 *         Map.of("name", "Widget A", "price", 29.99),
 *         Map.of("name", "Widget B", "price", 14.99)
 *     )
 * );
 *
 * String outputPath = MubanDocxEngine.builder()
 *     .template(new File("invoice.docx"))
 *     .data(data)
 *     .locale(Locale.forLanguageTag("pl-PL"))
 *     .outputDir("/tmp/output/")
 *     .outputFormat("pdf")
 *     .build()
 *     .generate();
 * }</pre>
 *
 * @see DocxExpressionEvaluator
 * @see DocxConditionalProcessor
 * @see DocxTableProcessor
 * @see DocxImageReplacer
 * @see DocxExporter
 */
public final class MubanDocxEngine {

    private static final Logger log = LoggerFactory.getLogger(MubanDocxEngine.class);

    /** Pattern to match ${placeholder} expressions */
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    /** Pattern for a simple key: Java identifier, optionally with dot-separated parts (e.g. address.city). */
    private static final Pattern SIMPLE_KEY_PATTERN = Pattern.compile("^[a-zA-Z_][a-zA-Z0-9_]*(\\.[a-zA-Z_][a-zA-Z0-9_]*)*$");

    private final WordprocessingMLPackage wordPackage;
    private final Map<String, Object> rawContext;
    private final Map<String, List<Map<String, Object>>> dataArrays;
    private final Map<String, String> formatMap;
    private final Locale locale;
    private final String outputDir;
    private final String outputFormat;
    private final File assetDir;
    private final PdfExportOptions pdfOptions;
    private final PdfSecurityCallback securityCallback;
    private final TxtExportOptions txtOptions;

    private MubanDocxEngine(WordprocessingMLPackage wordPackage,
                             Map<String, Object> rawContext,
                             Map<String, List<Map<String, Object>>> dataArrays,
                             Map<String, String> formatMap,
                             Locale locale,
                             String outputDir,
                             String outputFormat,
                             File assetDir,
                             PdfExportOptions pdfOptions,
                             PdfSecurityCallback securityCallback,
                             TxtExportOptions txtOptions) {
        this.wordPackage = wordPackage;
        this.rawContext = rawContext;
        this.dataArrays = dataArrays;
        this.formatMap = formatMap;
        this.locale = locale;
        this.outputDir = outputDir;
        this.outputFormat = outputFormat;
        this.assetDir = assetDir;
        this.pdfOptions = pdfOptions;
        this.securityCallback = securityCallback;
        this.txtOptions = txtOptions;
    }

    /**
     * Execute the full template processing pipeline and export.
     *
     * @return absolute path to the generated output file
     * @throws MubanDocxException if any processing step fails
     */
    public String generate() {
        long start = System.currentTimeMillis();

        // Step 0: Prepare document — merge split runs, clean up rsids/bookmarks
        try {
            VariablePrepare.prepare(wordPackage);
        } catch (Exception e) {
            throw new MubanDocxException("PREPARE_FAILED",
                    "Failed to prepare document for variable replacement: " + e.getMessage(), e);
        }

        // Step 1: Conditional blocks (body + headers/footers)
        processConditionalBlocks();

        // Step 2: Placeholder replacement (with table row replication)
        processContent(wordPackage.getMainDocumentPart().getContent());
        processHeadersAndFooters();

        // Step 3: Image replacement
        if (assetDir != null) {
            DocxImageReplacer.replaceImages(wordPackage, rawContext, assetDir, locale);
        }

        // Step 4: Export
        String result = DocxExporter.exportDocument(
                wordPackage, outputFormat, outputDir, pdfOptions, securityCallback, txtOptions);

        long elapsed = System.currentTimeMillis() - start;
        log.info("Template generation completed in {} ms → {}", elapsed, outputFormat.toUpperCase());

        return result;
    }

    /**
     * Process the template in-memory without exporting.
     *
     * <p>Runs the content processing pipeline (conditionals, placeholders,
     * table row replication, images) and returns the modified
     * {@link WordprocessingMLPackage}. The caller is responsible
     * for saving or further processing the package.
     *
     * <p><b>Note:</b> Unlike {@link #generate()}, this method does <em>not</em>
     * call {@link VariablePrepare#prepare(WordprocessingMLPackage)}. For DOCX
     * files loaded from disk (which may contain split runs, rsids and bookmarks),
     * call {@code VariablePrepare.prepare(wordPackage)} before invoking this method.
     *
     * @return the processed DOCX package
     */
    public WordprocessingMLPackage process() {
        processConditionalBlocks();
        processContent(wordPackage.getMainDocumentPart().getContent());
        processHeadersAndFooters();
        if (assetDir != null) {
            DocxImageReplacer.replaceImages(wordPackage, rawContext, assetDir, locale);
        }
        return wordPackage;
    }

    // ==================== PIPELINE STEPS ====================

    private void processConditionalBlocks() {
        int bodyBlocks = DocxConditionalProcessor.processConditionals(
                wordPackage.getMainDocumentPart().getContent(), rawContext, locale);

        int hfBlocks = 0;
        try {
            if (wordPackage.getDocumentModel() != null) {
                for (var section : wordPackage.getDocumentModel().getSections()) {
                    var hfp = section.getHeaderFooterPolicy();
                    if (hfp == null) continue;
                    hfBlocks += processConditionalInPart(hfp.getDefaultHeader());
                    hfBlocks += processConditionalInPart(hfp.getDefaultFooter());
                    hfBlocks += processConditionalInPart(hfp.getFirstHeader());
                    hfBlocks += processConditionalInPart(hfp.getFirstFooter());
                    hfBlocks += processConditionalInPart(hfp.getEvenHeader());
                    hfBlocks += processConditionalInPart(hfp.getEvenFooter());
                }
            }
        } catch (Exception e) {
            log.debug("Could not process headers/footers for conditionals: {}", e.getMessage());
        }

        if (bodyBlocks + hfBlocks > 0) {
            log.debug("Processed {} conditional block(s) (body: {}, headers/footers: {})",
                    bodyBlocks + hfBlocks, bodyBlocks, hfBlocks);
        }
    }

    @SuppressWarnings("unchecked")
    private int processConditionalInPart(Object part) {
        if (part == null) return 0;
        if (part instanceof org.docx4j.openpackaging.parts.JaxbXmlPart<?> jaxbPart) {
            try {
                Object content = jaxbPart.getContents();
                if (content instanceof ContentAccessor accessor) {
                    return DocxConditionalProcessor.processConditionals(
                            accessor.getContent(), rawContext, locale);
                }
            } catch (Exception e) {
                log.debug("Error processing conditional blocks in header/footer: {}", e.getMessage());
            }
        }
        return 0;
    }

    private void processHeadersAndFooters() {
        try {
            if (wordPackage.getDocumentModel() != null) {
                for (var section : wordPackage.getDocumentModel().getSections()) {
                    var hfp = section.getHeaderFooterPolicy();
                    if (hfp == null) continue;
                    processPartContent(hfp.getDefaultHeader());
                    processPartContent(hfp.getDefaultFooter());
                    processPartContent(hfp.getFirstHeader());
                    processPartContent(hfp.getFirstFooter());
                    processPartContent(hfp.getEvenHeader());
                    processPartContent(hfp.getEvenFooter());
                }
            }
        } catch (Exception e) {
            log.debug("Could not process headers/footers for placeholders: {}", e.getMessage());
        }
    }

    private void processPartContent(Object part) {
        if (part == null) return;
        if (part instanceof org.docx4j.openpackaging.parts.JaxbXmlPart<?> jaxbPart) {
            try {
                Object content = jaxbPart.getContents();
                if (content instanceof ContentAccessor accessor) {
                    processContent(accessor.getContent());
                }
            } catch (Exception e) {
                log.debug("Error processing header/footer content: {}", e.getMessage());
            }
        }
    }

    /**
     * Process a content list: tables first (with row replication), then paragraphs.
     */
    private void processContent(List<Object> content) {
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);

            if (unwrapped instanceof Tbl table) {
                DocxTableProcessor.processTable(table, rawContext, dataArrays, locale,
                        (c, ctx, arrays, loc) -> processContent(c));
            } else if (unwrapped instanceof P paragraph) {
                replacePlaceholders(paragraph);
            } else if (unwrapped instanceof ContentAccessor accessor) {
                processContent(accessor.getContent());
            }
        }
    }

    /**
     * Replace ${...} placeholders in a single paragraph.
     */
    private void replacePlaceholders(P paragraph) {
        // Merge split runs so each ${...} is within a single run
        VariablePrepare.joinupRuns(paragraph);
        // Normalize runs containing w:br (soft line breaks) into separate runs
        DocxXmlUtils.splitRunsAtBreaks(paragraph);

        for (Object obj : paragraph.getContent()) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (!(unwrapped instanceof R run)) continue;

            String text = DocxXmlUtils.getRunText(run);
            if (!text.contains("${")) continue;

            Matcher matcher = PLACEHOLDER_PATTERN.matcher(text);
            StringBuilder result = new StringBuilder();
            while (matcher.find()) {
                String body = matcher.group(1).trim();
                String replacement = DocxExpressionEvaluator.evaluate(body, rawContext, locale);

                // Apply format pattern for simple key placeholders only
                if (!formatMap.isEmpty() && isSimpleKey(body) && formatMap.containsKey(body)) {
                    Object rawValue = rawContext.get(body);
                    if (rawValue != null) {
                        String formatted = applyFormat(rawValue, formatMap.get(body), locale);
                        if (formatted != null) {
                            replacement = formatted;
                        }
                    }
                }

                matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
            }
            matcher.appendTail(result);
            DocxXmlUtils.setRunText(run, result.toString());
        }
    }

    // ==================== STATIC UTILITIES ====================

    /**
     * Extract placeholder keys from a DOCX template (for discovery/validation).
     *
     * @param wordPackage the loaded DOCX template
     * @return set of placeholder keys found in text and image alt-text
     */
    public static Set<String> extractImageKeys(WordprocessingMLPackage wordPackage) {
        return DocxImageReplacer.extractImageKeys(wordPackage);
    }

    /**
     * Load a DOCX template from a file.
     *
     * @param file the DOCX file
     * @return the loaded package
     * @throws MubanDocxException if loading fails
     */
    public static WordprocessingMLPackage load(File file) {
        try {
            return WordprocessingMLPackage.load(file);
        } catch (Exception e) {
            throw new MubanDocxException("LOAD_FAILED",
                    "Failed to load DOCX template: " + e.getMessage(), e);
        }
    }

    /**
     * Load a DOCX template from an input stream.
     *
     * @param inputStream the DOCX input stream
     * @return the loaded package
     * @throws MubanDocxException if loading fails
     */
    public static WordprocessingMLPackage load(InputStream inputStream) {
        try {
            return WordprocessingMLPackage.load(inputStream);
        } catch (Exception e) {
            throw new MubanDocxException("LOAD_FAILED",
                    "Failed to load DOCX template from stream: " + e.getMessage(), e);
        }
    }

    // ==================== FORMAT UTILITIES ====================

    /**
     * Check if a placeholder body is a simple key lookup (not a SpEL expression).
     * Simple keys are identifiers like {@code amount}, {@code address.city}, {@code item_count}.
     * Anything with operators, method calls, ternary, or brackets is NOT a simple key.
     */
    public static boolean isSimpleKey(String body) {
        return SIMPLE_KEY_PATTERN.matcher(body).matches();
    }

    /**
     * Apply a format pattern to a raw value.
     *
     * <p>For {@link Number} values, uses {@link DecimalFormat} with the given pattern.
     * For {@link TemporalAccessor} values (LocalDate, LocalDateTime), uses {@link DateTimeFormatter}.
     * For {@link java.util.Date} values, uses {@link java.text.SimpleDateFormat}.
     *
     * @param rawValue the typed value from the context
     * @param pattern  the format pattern string
     * @param locale   the document locale (may be null)
     * @return formatted string, or {@code null} if the value type is not formattable or pattern is invalid
     */
    public static String applyFormat(Object rawValue, String pattern, Locale locale) {
        Locale effectiveLocale = locale != null ? locale : Locale.getDefault();
        try {
            if (rawValue instanceof Number number) {
                DecimalFormat df = new DecimalFormat(pattern, DecimalFormatSymbols.getInstance(effectiveLocale));
                return df.format(number);
            }
            if (rawValue instanceof TemporalAccessor temporal) {
                DateTimeFormatter dtf = DateTimeFormatter.ofPattern(pattern, effectiveLocale);
                return dtf.format(temporal);
            }
            if (rawValue instanceof java.util.Date date) {
                java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat(pattern, effectiveLocale);
                return sdf.format(date);
            }
            return null; // not a formattable type — keep evaluator's output
        } catch (IllegalArgumentException e) {
            log.warn("Format pattern '{}' failed for value '{}' ({}): {}",
                    pattern, rawValue, rawValue.getClass().getSimpleName(), e.getMessage());
            return null; // graceful fallback
        }
    }

    // ==================== BUILDER ====================

    /**
     * Create a new builder for the template engine.
     */
    public static Builder builder() {
        return new Builder();
    }

    /**
     * Fluent builder for {@link MubanDocxEngine}.
     */
    public static class Builder {
        private WordprocessingMLPackage wordPackage;
        private Map<String, String> parameters;
        private Map<String, Object> data;
        private Map<String, Object> rawContext;
        private Locale locale;
        private String outputDir = System.getProperty("java.io.tmpdir");
        private String outputFormat = "docx";
        private File assetDir;
        private PdfExportOptions pdfOptions;
        private PdfSecurityCallback securityCallback;
        private TxtExportOptions txtOptions;
        private Map<String, String> formatMap;

        /**
         * Set the template from a pre-loaded package.
         */
        public Builder template(WordprocessingMLPackage pkg) {
            this.wordPackage = pkg;
            return this;
        }

        /**
         * Load the template from a file.
         */
        public Builder template(File file) {
            this.wordPackage = MubanDocxEngine.load(file);
            return this;
        }

        /**
         * Load the template from an input stream.
         */
        public Builder template(InputStream inputStream) {
            this.wordPackage = MubanDocxEngine.load(inputStream);
            return this;
        }

        /**
         * Set simple string parameters (added to context before data).
         */
        public Builder parameters(Map<String, String> params) {
            this.parameters = params;
            return this;
        }

        /**
         * Set the data map (may contain nested maps, arrays, numbers, booleans).
         */
        public Builder data(Map<String, Object> data) {
            this.data = data;
            return this;
        }

        /**
         * Set a pre-built raw context directly (bypasses parameters + data merging).
         */
        public Builder rawContext(Map<String, Object> ctx) {
            this.rawContext = ctx;
            return this;
        }

        /**
         * Set the locale for number/date formatting in expressions.
         */
        public Builder locale(Locale locale) {
            this.locale = locale;
            return this;
        }

        /**
         * Set the output directory (defaults to system temp dir).
         */
        public Builder outputDir(String dir) {
            this.outputDir = dir;
            return this;
        }

        /**
         * Set the output format: "docx", "pdf", "html", or "txt" (defaults to "docx").
         */
        public Builder outputFormat(String format) {
            this.outputFormat = format;
            return this;
        }

        /**
         * Set the asset directory for image resolution (extracted template root).
         */
        public Builder assetDir(File dir) {
            this.assetDir = dir;
            return this;
        }

        /**
         * Set PDF export options (security, encryption).
         */
        public Builder pdfOptions(PdfExportOptions opts) {
            this.pdfOptions = opts;
            return this;
        }

        /**
         * Set the callback for applying PDF security.
         */
        public Builder pdfSecurityCallback(PdfSecurityCallback cb) {
            this.securityCallback = cb;
            return this;
        }

        /**
         * Set TXT export options (line separator).
         */
        public Builder txtOptions(TxtExportOptions opts) {
            this.txtOptions = opts;
            return this;
        }

        /**
         * Set the format map for post-evaluation formatting.
         * Maps placeholder names to format patterns (e.g. {@code #,##0.00}, {@code dd.MM.yyyy}).
         */
        public Builder formatMap(Map<String, String> fm) {
            this.formatMap = fm;
            return this;
        }

        /**
         * Build the engine instance.
         *
         * @throws IllegalStateException if no template is set
         */
        public MubanDocxEngine build() {
            if (wordPackage == null) {
                throw new IllegalStateException("Template must be set via template()");
            }

            // Build context: explicit rawContext wins over parameters+data merging
            Map<String, Object> ctx;
            if (rawContext != null) {
                ctx = rawContext;
            } else {
                ctx = DocxContextBuilder.buildRawContext(parameters, data);
            }

            Map<String, List<Map<String, Object>>> arrays = DocxContextBuilder.extractDataArrays(
                    data != null ? data : ctx);

            return new MubanDocxEngine(
                    wordPackage, ctx, arrays,
                    formatMap != null ? formatMap : Collections.emptyMap(),
                    locale, outputDir, outputFormat, assetDir,
                    pdfOptions, securityCallback, txtOptions);
        }
    }
}
