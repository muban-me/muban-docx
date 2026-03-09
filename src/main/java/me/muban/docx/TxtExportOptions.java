package me.muban.docx;

/**
 * Plain-text export options for DOCX → TXT conversion.
 *
 * @param lineSeparator    string inserted between paragraphs; if {@code null},
 *                         {@link System#lineSeparator()} is used
 * @param trimLineRight    whether to strip trailing whitespace from each paragraph
 * @param pageWidthInChars soft word-wrap width; if {@code null}, no wrapping is applied
 */
public record TxtExportOptions(String lineSeparator, boolean trimLineRight, Integer pageWidthInChars) {

    /** Default options — system line separator, no trimming, no wrapping. */
    public TxtExportOptions() {
        this(null, false, null);
    }

    /** Convenience — custom separator only, no trimming, no wrapping. */
    public TxtExportOptions(String lineSeparator) {
        this(lineSeparator, false, null);
    }

    /** Convenience — custom separator and trimming, no wrapping. */
    public TxtExportOptions(String lineSeparator, boolean trimLineRight) {
        this(lineSeparator, trimLineRight, null);
    }

    /** Resolved separator: returns the configured value or the system default. */
    public String resolvedLineSeparator() {
        return lineSeparator != null ? lineSeparator : System.lineSeparator();
    }
}
