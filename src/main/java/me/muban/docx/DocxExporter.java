package me.muban.docx;

import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.UUID;

/**
 * Exports a filled DOCX document to the requested output format (DOCX or PDF).
 *
 * <p>For PDF output, resolves table conditional formatting before export (since the
 * XSL-FO pipeline doesn't handle table style properties), sets up font mapping,
 * and optionally applies PDF security settings via a caller-provided
 * {@link PdfSecurityCallback}.
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
     * @param format         output format: {@code "docx"} or {@code "pdf"}
     * @param outputDir      directory to write the output file to (created if needed)
     * @param pdfOptions     optional PDF security options (only used for PDF output); may be null
     * @param securityCallback optional callback to apply PDF encryption; may be null
     * @return absolute path to the generated file
     * @throws MubanDocxException if export fails
     * @throws UnsupportedOutputFormatException if format is not {@code "docx"} or {@code "pdf"}
     */
    public static String exportDocument(WordprocessingMLPackage wordPackage,
                                         String format,
                                         String outputDir,
                                         PdfExportOptions pdfOptions,
                                         PdfSecurityCallback securityCallback) {
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
     * Export to DOCX only (convenience overload).
     */
    public static String exportDocx(WordprocessingMLPackage wordPackage, String outputDir) {
        return exportDocument(wordPackage, "docx", outputDir, null, null);
    }

    /**
     * Export to PDF (convenience overload without security).
     */
    public static String exportPdf(WordprocessingMLPackage wordPackage, String outputDir) {
        return exportDocument(wordPackage, "pdf", outputDir, null, null);
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
}
