package me.muban.docx;

/**
 * HTML export options for DOCX &#8594; HTML conversion.
 *
 * @param imageMode how images are represented in the output; {@code null}
 *                  resolves to {@link #DEFAULT_IMAGE_MODE}
 */
public record HtmlExportOptions(HtmlImageMode imageMode) {

    /** Mode used when the caller does not specify one. */
    public static final HtmlImageMode DEFAULT_IMAGE_MODE = HtmlImageMode.FILES;

    public HtmlExportOptions {
        if (imageMode == null) {
            imageMode = DEFAULT_IMAGE_MODE;
        }
    }

    /** @return options using the default image mode ({@link #DEFAULT_IMAGE_MODE}) */
    public static HtmlExportOptions defaults() {
        return new HtmlExportOptions(DEFAULT_IMAGE_MODE);
    }
}
