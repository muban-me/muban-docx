package me.muban.docx;

/**
 * Options for PNG export.
 *
 * <p>The zoom ratio multiplies the native page dimensions (in points at 72 DPI)
 * to produce the output raster size. For example, a zoom ratio of {@code 2.0}
 * produces an image twice as wide and twice as tall as the native page.</p>
 *
 * @param zoomRatio zoom ratio for the output raster; must be a positive, finite number
 */
public record PngExportOptions(float zoomRatio) {

    /** Default zoom ratio: native page size (~72 DPI). */
    public static final float DEFAULT_ZOOM_RATIO = 1.0f;

    public PngExportOptions {
        if (!Float.isFinite(zoomRatio) || zoomRatio <= 0f) {
            throw new IllegalArgumentException("zoomRatio must be a positive number: " + zoomRatio);
        }
    }

    /**
     * @return options using the default zoom ratio ({@value #DEFAULT_ZOOM_RATIO})
     */
    public static PngExportOptions defaults() {
        return new PngExportOptions(DEFAULT_ZOOM_RATIO);
    }
}
