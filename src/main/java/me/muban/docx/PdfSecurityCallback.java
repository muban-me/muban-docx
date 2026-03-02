package me.muban.docx;

/**
 * Callback interface for applying PDF security settings (encryption, passwords, permissions)
 * to a generated PDF file.
 *
 * <p>The muban-docx library does not bundle a specific PDF security implementation.
 * Consumers provide their own via this callback, which is invoked after PDF export
 * when security options are present.
 *
 * <p>Example implementation using PDFBox:
 * <pre>{@code
 * PdfSecurityCallback security = (pdfPath, options) -> {
 *     try (PDDocument doc = Loader.loadPDF(new File(pdfPath))) {
 *         AccessPermission ap = new AccessPermission();
 *         ap.setCanPrint(options.canPrint());
 *         // ... set other permissions
 *         StandardProtectionPolicy policy = new StandardProtectionPolicy(
 *             options.ownerPassword(), options.userPassword(), ap);
 *         doc.protect(policy);
 *         doc.save(pdfPath);
 *     }
 * };
 * }</pre>
 *
 * @see PdfExportOptions
 */
@FunctionalInterface
public interface PdfSecurityCallback {

    /**
     * Apply security settings to a PDF file on disk.
     *
     * @param pdfFilePath absolute path to the PDF file to protect
     * @param options     the security options to apply
     * @throws Exception if security application fails
     */
    void applySecurity(String pdfFilePath, PdfExportOptions options) throws Exception;
}
