package me.muban.docx;

/**
 * How images are represented in DOCX &#8594; HTML output.
 *
 * @see HtmlExportOptions
 */
public enum HtmlImageMode {

    /**
     * Image files are written next to {@code index.html} in the
     * {@code index.html_files/} directory and referenced by relative path
     * (default, matches the JasperReports HTML convention).
     */
    FILES,

    /**
     * No image files are written and image references are left empty.
     * Callers that need an image-free HTML document must strip the image
     * elements themselves — the Muban service does this in its HTML
     * post-processing pipeline.
     */
    NONE,

    /**
     * Images are inlined as RFC 2397 data URIs
     * ({@code src="data:image/png;base64,..."}). Convenient for a single-file
     * HTML artifact, but poorly supported by email clients (Gmail strips
     * {@code data:} URIs) and it inflates the HTML by ~33%.
     */
    DATA_URI,

    /**
     * Images are referenced as {@code cid:} URIs (RFC 2392) for a
     * {@code multipart/related} email body. The image files are still written
     * to {@code index.html_files/} so the sender can attach them, and
     * {@code images.json} maps every Content-ID to its file.
     */
    CID;

    /**
     * Parses a mode name case-insensitively.
     *
     * @param value mode name, e.g. {@code "data_uri"}
     * @return the matching mode
     * @throws IllegalArgumentException if the value does not name a mode
     */
    public static HtmlImageMode fromString(String value) {
        for (HtmlImageMode mode : values()) {
            if (mode.name().equalsIgnoreCase(value)) {
                return mode;
            }
        }
        throw new IllegalArgumentException("Unsupported image mode: " + value
                + " (expected one of FILES, NONE, DATA_URI, CID)");
    }
}
