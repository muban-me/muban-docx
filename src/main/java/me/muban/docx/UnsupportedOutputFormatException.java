package me.muban.docx;

/**
 * Thrown when the requested output format is not supported.
 *
 * <p>The muban-docx library supports DOCX (passthrough) and PDF (via XSL-FO) output.
 * Any other format will cause this exception.
 */
public class UnsupportedOutputFormatException extends MubanDocxException {

    private final String requestedFormat;

    /**
     * @param format the unsupported format that was requested
     */
    public UnsupportedOutputFormatException(String format) {
        super("UNSUPPORTED_FORMAT",
              String.format("Output format '%s' is not supported. Supported formats: docx, pdf", format),
              null);
        this.requestedFormat = format;
    }

    /**
     * @return the format string that caused this exception
     */
    public String getRequestedFormat() {
        return requestedFormat;
    }
}
