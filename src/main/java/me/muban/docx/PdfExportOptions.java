package me.muban.docx;

/**
 * PDF export options — security, encryption, and permissions.
 *
 * <p>This record replaces the service-layer DTO with a framework-independent value object.
 * All fields are optional; security settings are only applied when at least one of
 * {@code userPassword} or {@code ownerPassword} is non-empty.
 *
 * @param userPassword          password required to open the PDF (may be null)
 * @param ownerPassword         password required to change security settings (may be null)
 * @param canPrint              allow printing (default true)
 * @param canPrintHighQuality   allow high-quality printing (default true)
 * @param canModify             allow content modification (default false)
 * @param canCopy               allow copying text/graphics (default true)
 * @param canFillForms          allow filling forms (default true)
 * @param canAnnotate           allow annotations (default true)
 * @param canAssemble           allow page assembly (default false)
 * @param encryptionKeyLength   128 or 256 bits (default 128)
 */
public record PdfExportOptions(
        String userPassword,
        String ownerPassword,
        boolean canPrint,
        boolean canPrintHighQuality,
        boolean canModify,
        boolean canCopy,
        boolean canFillForms,
        boolean canAnnotate,
        boolean canAssemble,
        int encryptionKeyLength
) {

    /**
     * Create options with sensible defaults — security disabled.
     */
    public PdfExportOptions() {
        this(null, null, true, true, false, true, true, true, false, 128);
    }

    /**
     * Create options with just user and owner passwords (all defaults).
     *
     * @param userPassword  password to open the PDF
     * @param ownerPassword password to change security settings
     */
    public PdfExportOptions(String userPassword, String ownerPassword) {
        this(userPassword, ownerPassword, true, true, false, true, true, true, false, 128);
    }

    /**
     * @return true if any security setting (password) is configured
     */
    public boolean hasSecuritySettings() {
        return (userPassword != null && !userPassword.isEmpty())
                || (ownerPassword != null && !ownerPassword.isEmpty());
    }

    /**
     * Builder for creating {@link PdfExportOptions} with fine-grained control.
     */
    public static Builder builder() {
        return new Builder();
    }

    /**
     * Mutable builder for {@link PdfExportOptions}.
     */
    public static class Builder {
        private String userPassword;
        private String ownerPassword;
        private boolean canPrint = true;
        private boolean canPrintHighQuality = true;
        private boolean canModify = false;
        private boolean canCopy = true;
        private boolean canFillForms = true;
        private boolean canAnnotate = true;
        private boolean canAssemble = false;
        private int encryptionKeyLength = 128;

        public Builder userPassword(String val) { this.userPassword = val; return this; }
        public Builder ownerPassword(String val) { this.ownerPassword = val; return this; }
        public Builder canPrint(boolean val) { this.canPrint = val; return this; }
        public Builder canPrintHighQuality(boolean val) { this.canPrintHighQuality = val; return this; }
        public Builder canModify(boolean val) { this.canModify = val; return this; }
        public Builder canCopy(boolean val) { this.canCopy = val; return this; }
        public Builder canFillForms(boolean val) { this.canFillForms = val; return this; }
        public Builder canAnnotate(boolean val) { this.canAnnotate = val; return this; }
        public Builder canAssemble(boolean val) { this.canAssemble = val; return this; }
        public Builder encryptionKeyLength(int val) { this.encryptionKeyLength = val; return this; }

        public PdfExportOptions build() {
            return new PdfExportOptions(userPassword, ownerPassword,
                    canPrint, canPrintHighQuality, canModify, canCopy,
                    canFillForms, canAnnotate, canAssemble, encryptionKeyLength);
        }
    }
}
