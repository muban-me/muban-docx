package me.muban.docx;

import org.docx4j.Docx4J;
import org.docx4j.TextUtils;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.convert.out.html.HTMLConversionImageHandler;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Comparator;
import java.util.List;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Exports a filled DOCX document to the requested output format (DOCX, PDF, HTML, or TXT).
 *
 * <p>For PDF output, resolves table conditional formatting before export (since the
 * XSL-FO pipeline doesn't handle table style properties), sets up font mapping,
 * and optionally applies PDF security settings via a caller-provided
 * {@link PdfSecurityCallback}.
 *
 * <p>For HTML output, exports using docx4j's XSLT-based HTML exporter. Images are
 * saved to an {@code index.html_files/} subdirectory (matching the JasperReports
 * HTML convention for S2S compatibility). The resulting directory is packaged as a
 * ZIP archive containing {@code index.html} and its assets.
 *
 * <p><b>Font mapper ordering:</b> {@code setFontMapper()} must be called before
 * {@code FOSettings.setOpcPackage()}, because {@code setOpcPackage()} uses the mapper
 * to build FOP font config which maps bold/italic variants.
 */
public class DocxExporter {

    private static final Logger log = LoggerFactory.getLogger(DocxExporter.class);

    private DocxExporter() {}

    /**
     * Export the filled document to the requested format.
     *
     * @param wordPackage    the filled DOCX document
     * @param format         output format: {@code "docx"}, {@code "pdf"}, {@code "html"}, or {@code "txt"}
     * @param outputDir      directory to write the output file to (created if needed)
     * @param pdfOptions     optional PDF security options (only used for PDF output); may be null
     * @param securityCallback optional callback to apply PDF encryption; may be null
     * @param txtOptions     optional TXT export options (only used for TXT output); may be null
     * @return absolute path to the generated file (ZIP archive for HTML format)
     * @throws MubanDocxException if export fails
     * @throws UnsupportedOutputFormatException if format is not supported
     */
    public static String exportDocument(WordprocessingMLPackage wordPackage,
                                         String format,
                                         String outputDir,
                                         PdfExportOptions pdfOptions,
                                         PdfSecurityCallback securityCallback,
                                         TxtExportOptions txtOptions) {
        try {
            Files.createDirectories(Paths.get(outputDir));
        } catch (Exception e) {
            throw new MubanDocxException("EXPORT_FAILED",
                    "Cannot create output directory: " + outputDir, e);
        }

        String filename = UUID.randomUUID().toString();
        String outputPath;

        switch (format.toLowerCase()) {
            case "docx" -> {
                outputPath = outputDir + File.separator + filename + ".docx";
                try {
                    wordPackage.save(new File(outputPath));
                } catch (Exception e) {
                    throw new MubanDocxException("EXPORT_FAILED",
                            "DOCX save failed: " + e.getMessage(), e);
                }
            }
            case "pdf" -> {
                outputPath = outputDir + File.separator + filename + ".pdf";
                exportToPdf(wordPackage, outputPath, pdfOptions, securityCallback);
            }
            case "html" -> {
                outputPath = outputDir + File.separator + filename + ".zip";
                exportToHtml(wordPackage, outputDir, filename, outputPath);
            }
            case "txt" -> {
                outputPath = outputDir + File.separator + filename + ".txt";
                exportToTxt(wordPackage, outputPath, txtOptions);
            }
            default -> throw new UnsupportedOutputFormatException(format);
        }

        // Verify output file
        File outputFile = new File(outputPath);
        if (!outputFile.exists() || outputFile.length() == 0) {
            throw new MubanDocxException("EXPORT_FAILED",
                    format.toUpperCase() + " export failed — output file is empty or not created", null);
        }

        log.debug("{} export completed. File size: {} bytes", format.toUpperCase(), outputFile.length());
        return outputPath;
    }

    /**
     * Backward-compatible overload without TXT options.
     */
    public static String exportDocument(WordprocessingMLPackage wordPackage,
                                         String format,
                                         String outputDir,
                                         PdfExportOptions pdfOptions,
                                         PdfSecurityCallback securityCallback) {
        return exportDocument(wordPackage, format, outputDir, pdfOptions, securityCallback, null);
    }

    /**
     * Export to DOCX only (convenience overload).
     */
    public static String exportDocx(WordprocessingMLPackage wordPackage, String outputDir) {
        return exportDocument(wordPackage, "docx", outputDir, null, null, null);
    }

    /**
     * Export to PDF (convenience overload without security).
     */
    public static String exportPdf(WordprocessingMLPackage wordPackage, String outputDir) {
        return exportDocument(wordPackage, "pdf", outputDir, null, null, null);
    }

    /**
     * Export to HTML as ZIP archive (convenience overload).
     */
    public static String exportHtml(WordprocessingMLPackage wordPackage, String outputDir) {
        return exportDocument(wordPackage, "html", outputDir, null, null, null);
    }

    /**
     * Export to plain text (convenience overload).
     */
    public static String exportTxt(WordprocessingMLPackage wordPackage, String outputDir) {
        return exportDocument(wordPackage, "txt", outputDir, null, null, null);
    }

    /**
     * Export to plain text with custom options (convenience overload).
     */
    public static String exportTxt(WordprocessingMLPackage wordPackage, String outputDir,
                                    TxtExportOptions txtOptions) {
        return exportDocument(wordPackage, "txt", outputDir, null, null, txtOptions);
    }

    /**
     * Internal TXT export — walks all {@code w:p} paragraphs (including those
     * nested inside tables) and writes each paragraph's text separated by the
     * configured line separator (or system default).
     */
    private static void exportToTxt(WordprocessingMLPackage wordPackage, String outputPath,
                                     TxtExportOptions txtOptions) {
        String separator = txtOptions != null
                ? txtOptions.resolvedLineSeparator()
                : System.lineSeparator();
        boolean trim = txtOptions != null && txtOptions.trimLineRight();
        Integer wrapWidth = txtOptions != null ? txtOptions.pageWidthInChars() : null;
        try (Writer writer = new BufferedWriter(
                new OutputStreamWriter(new FileOutputStream(outputPath), StandardCharsets.UTF_8))) {
            List<Object> paragraphs = wordPackage.getMainDocumentPart()
                    .getJAXBNodesViaXPath("//w:p", true);
            for (int i = 0; i < paragraphs.size(); i++) {
                if (i > 0) writer.write(separator);
                String text = TextUtils.getText(paragraphs.get(i));
                if (text == null) text = "";
                if (trim) text = text.stripTrailing();
                if (wrapWidth != null && wrapWidth > 0) {
                    writer.write(softWrap(text, wrapWidth, separator));
                } else {
                    writer.write(text);
                }
            }
        } catch (Exception e) {
            throw new MubanDocxException("EXPORT_FAILED",
                    "TXT export failed: " + e.getMessage(), e);
        }
    }

    /**
     * Soft word-wrap: breaks {@code text} at spaces so no line exceeds
     * {@code width} characters.  Words longer than {@code width} are
     * kept intact on their own line (never broken).
     */
    static String softWrap(String text, int width, String lineSep) {
        if (text.length() <= width) return text;
        StringBuilder sb = new StringBuilder(text.length() + 16);
        int lineLen = 0;
        for (String word : text.split(" ", -1)) {
            if (lineLen == 0) {
                sb.append(word);
                lineLen = word.length();
            } else if (lineLen + 1 + word.length() <= width) {
                sb.append(' ').append(word);
                lineLen += 1 + word.length();
            } else {
                sb.append(lineSep).append(word);
                lineLen = word.length();
            }
        }
        return sb.toString();
    }

    /**
     * Internal PDF export via docx4j's XSL-FO pipeline.
     */
    private static void exportToPdf(WordprocessingMLPackage wordPackage,
                                     String outputPath,
                                     PdfExportOptions pdfOptions,
                                     PdfSecurityCallback securityCallback) {
        // Resolve table conditional formatting (cnfStyle) before FO export
        DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

        // Font mapper MUST be set before FOSettings.setOpcPackage()
        try {
            wordPackage.setFontMapper(new IdentityPlusMapper());
        } catch (Exception e) {
            log.warn("Could not set up font mapper for PDF export, " +
                     "bold/italic may not render correctly: {}", e.getMessage());
        }

        FOSettings foSettings = Docx4J.createFOSettings();
        foSettings.setApacheFopMime("application/pdf");

        try (OutputStream os = new FileOutputStream(outputPath)) {
            foSettings.setOpcPackage(wordPackage);
            Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
        } catch (Exception e) {
            throw new MubanDocxException("EXPORT_FAILED",
                    "PDF export failed: " + e.getMessage(), e);
        }

        // Post-process: apply PDF security settings if requested
        if (pdfOptions != null && pdfOptions.hasSecuritySettings() && securityCallback != null) {
            try {
                log.info("Applying PDF security settings");
                securityCallback.applySecurity(outputPath, pdfOptions);
                log.info("PDF security settings applied successfully");
            } catch (Exception e) {
                throw new MubanDocxException("PDF_SECURITY_FAILED",
                        "Failed to apply PDF security: " + e.getMessage(), e);
            }
        }
    }

    /**
     * Internal HTML export via docx4j's XSLT-based HTML exporter.
     *
     * <p>Creates a temporary directory with {@code index.html} and an
     * {@code index.html_files/} subdirectory for images, then packages
     * everything into a ZIP archive.
     */
    private static void exportToHtml(WordprocessingMLPackage wordPackage,
                                      String outputDir,
                                      String dirName,
                                      String zipOutputPath) {
        // Resolve table conditional formatting (borders, shading, text formatting)
        // before HTML export — same as PDF pipeline — because the XSLT exporter
        // doesn't resolve Word's tblStylePr/cnfStyle references.
        DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

        // Set up font mapper so the HTML exporter can resolve font families
        try {
            wordPackage.setFontMapper(new IdentityPlusMapper());
        } catch (Exception e) {
            log.warn("Could not set up font mapper for HTML export: {}", e.getMessage());
        }

        Path htmlDir = Paths.get(outputDir, dirName);
        try {
            Files.createDirectories(htmlDir);

            String imageDirPath = htmlDir.resolve("index.html_files").toString();
            Files.createDirectories(Paths.get(imageDirPath));

            // Export HTML using docx4j
            HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
            htmlSettings.setOpcPackage(wordPackage);
            htmlSettings.setFontMapper(wordPackage.getFontMapper());
            htmlSettings.setImageDirPath(imageDirPath);
            htmlSettings.setImageTargetUri("index.html_files/");
            htmlSettings.setImageHandler(
                    new HTMLConversionImageHandler(imageDirPath, "index.html_files/", true));

            Path indexHtml = htmlDir.resolve("index.html");
            try (OutputStream os = new FileOutputStream(indexHtml.toFile())) {
                Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
            }

            // Package into ZIP archive
            zipDirectory(htmlDir, zipOutputPath);
        } catch (Exception e) {
            throw new MubanDocxException("EXPORT_FAILED",
                    "HTML export failed: " + e.getMessage(), e);
        } finally {
            // Clean up temporary HTML directory
            deleteDirectoryQuietly(htmlDir);
        }
    }

    /**
     * Create a ZIP archive from a directory. Entries are stored with the
     * directory name as the root folder inside the ZIP.
     */
    private static void zipDirectory(Path sourceDir, String zipFilePath) throws IOException {
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFilePath))) {
            String rootName = sourceDir.getFileName().toString();
            Files.walkFileTree(sourceDir, new SimpleFileVisitor<>() {
                @Override
                public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs)
                        throws IOException {
                    String entryName = rootName + "/" + sourceDir.relativize(dir);
                    if (!entryName.endsWith("/")) entryName += "/";
                    zos.putNextEntry(new ZipEntry(entryName));
                    zos.closeEntry();
                    return FileVisitResult.CONTINUE;
                }

                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                        throws IOException {
                    String entryName = rootName + "/" + sourceDir.relativize(file);
                    zos.putNextEntry(new ZipEntry(entryName));
                    Files.copy(file, zos);
                    zos.closeEntry();
                    return FileVisitResult.CONTINUE;
                }
            });
        }
    }

    private static void deleteDirectoryQuietly(Path dir) {
        try {
            if (Files.exists(dir)) {
                Files.walkFileTree(dir, new SimpleFileVisitor<>() {
                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                            throws IOException {
                        Files.delete(file);
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult postVisitDirectory(Path d, IOException exc)
                            throws IOException {
                        Files.delete(d);
                        return FileVisitResult.CONTINUE;
                    }
                });
            }
        } catch (IOException e) {
            log.warn("Could not clean up temporary HTML directory {}: {}", dir, e.getMessage());
        }
    }
}
