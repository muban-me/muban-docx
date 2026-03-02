package me.muban.docx;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.InputStream;
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

    private final WordprocessingMLPackage wordPackage;
    private final Map<String, Object> rawContext;
    private final Map<String, List<Map<String, Object>>> dataArrays;
    private final Locale locale;
    private final String outputDir;
    private final String outputFormat;
    private final File assetDir;
    private final PdfExportOptions pdfOptions;
    private final PdfSecurityCallback securityCallback;

    private MubanDocxEngine(WordprocessingMLPackage wordPackage,
                             Map<String, Object> rawContext,
                             Map<String, List<Map<String, Object>>> dataArrays,
                             Locale locale,
                             String outputDir,
                             String outputFormat,
                             File assetDir,
                             PdfExportOptions pdfOptions,
                             PdfSecurityCallback securityCallback) {
        this.wordPackage = wordPackage;
        this.rawContext = rawContext;
        this.dataArrays = dataArrays;
        this.locale = locale;
        this.outputDir = outputDir;
        this.outputFormat = outputFormat;
        this.assetDir = assetDir;
        this.pdfOptions = pdfOptions;
        this.securityCallback = securityCallback;
    }

    /**
     * Execute the full template processing pipeline and export.
     *
     * @return absolute path to the generated output file
     * @throws MubanDocxException if any processing step fails
     */
    public String generate() {
        long start = System.currentTimeMillis();

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
                wordPackage, outputFormat, outputDir, pdfOptions, securityCallback);

        long elapsed = System.currentTimeMillis() - start;
        log.info("Template generation completed in {} ms → {}", elapsed, outputFormat.toUpperCase());

        return result;
    }

    /**
     * Process the template in-memory without exporting.
     *
     * <p>Runs steps 1–3 (conditionals, placeholders, images) and returns
     * the modified {@link WordprocessingMLPackage}. The caller is responsible
     * for saving or further processing the package.
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
         * Set the output format: "docx" or "pdf" (defaults to "docx").
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
                    wordPackage, ctx, arrays, locale,
                    outputDir, outputFormat, assetDir,
                    pdfOptions, securityCallback);
        }
    }
}
