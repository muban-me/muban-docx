package me.muban.docx;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.WriterException;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.QRCodeWriter;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.EnumMap;
import java.util.Locale;
import java.util.Map;

/**
 * Renders QR codes as PNG images for use with dynamic image placeholders.
 *
 * <p>Complements the image resolution cascade of {@link DocxImageReplacer}: a context value
 * (or an evaluated alt-text expression) starting with {@code qr:} is encoded into a PNG
 * QR code instead of being interpreted as a data URI, URL or path.</p>
 *
 * <p><b>Example</b> — image alt text {@code image:qr_code} with context value:</p>
 * <pre>{@code
 * Map.of("qr_code", "qr:https://www.warta.pl/obsluga/platnosci?nr=920046666381")
 * }</pre>
 *
 * <p>Rendering can be tuned per image with a query string appended to the alt text key,
 * e.g. {@code image:qr_code?ec=H&margin=4&bg=transparent}:</p>
 *
 * <table border="1">
 *   <caption>Supported options</caption>
 *   <tr><th>Option</th><th>Values</th><th>Default</th><th>Meaning</th></tr>
 *   <tr><td>{@code ec}</td><td>L, M, Q, H</td><td>M</td><td>error correction level</td></tr>
 *   <tr><td>{@code margin}</td><td>0-16</td><td>2</td><td>quiet zone, in modules</td></tr>
 *   <tr><td>{@code scale}</td><td>1-40</td><td>auto</td><td>pixels per module (disables auto-scaling)</td></tr>
 *   <tr><td>{@code minPx}</td><td>60-4000</td><td>300</td><td>minimum edge length when scaling automatically</td></tr>
 *   <tr><td>{@code bg}</td><td>{@code white}, {@code transparent}, {@code #RRGGBB}</td><td>white</td>
 *       <td>background colour; transparent keeps white RGB under the alpha channel</td></tr>
 * </table>
 *
 * <p>UTF-8 is used for the payload. The picture frame (size and position) is taken from the
 * DOCX template, so the PNG is generated with an oversampled module size and scaled down by
 * Word — which keeps module edges crisp.</p>
 *
 * <p>Encoding never throws on bad input: {@code null}, blank payloads and encoder failures
 * are logged and reported as {@code null}, which makes {@link DocxImageReplacer} keep the
 * original template image.</p>
 */
public final class QrCodeEncoder {

    /** Value prefix marking a context value as a QR payload to be rendered. */
    public static final String QR_PREFIX = "qr:";

    private static final Logger log = LoggerFactory.getLogger(QrCodeEncoder.class);

    private QrCodeEncoder() {}

    /**
     * Error correction level — mirrors the QR specification, independent of the encoder library.
     */
    public enum ErrorCorrection {
        /** ~7% recovery. */
        L,
        /** ~15% recovery — the usual choice for printed documents. */
        M,
        /** ~25% recovery. */
        Q,
        /** ~30% recovery — largest symbol for a given payload. */
        H
    }

    /**
     * Rendering options for {@link #toPng(String, QrOptions)}.
     *
     * @param errorCorrection error correction level, never {@code null}
     * @param marginModules   quiet zone in modules (0-16)
     * @param scale           pixels per module, or {@link #AUTO_SCALE} to derive it from {@code minPixels}
     * @param minPixels       minimum edge length in pixels when {@code scale} is automatic
     * @param background      {@code white}, {@code transparent} or {@code #RRGGBB}
     */
    public record QrOptions(ErrorCorrection errorCorrection,
                            int marginModules,
                            int scale,
                            int minPixels,
                            String background) {

        /** Value of {@link #scale()} meaning "derive the module size from {@code minPixels}". */
        public static final int AUTO_SCALE = 0;

        private static final String BACKGROUND_WHITE = "white";
        private static final String BACKGROUND_TRANSPARENT = "transparent";
        private static final int MIN_MARGIN = 0;
        private static final int MAX_MARGIN = 16;
        private static final int MIN_PIXELS_FLOOR = 60;
        private static final int MAX_PIXELS = 4000;
        private static final int INVALID = Integer.MIN_VALUE;

        /**
         * Default options: error correction {@code M}, 2-module quiet zone, automatic scaling
         * to at least 300 px, opaque white background.
         *
         * @return default options, never {@code null}
         */
        public static QrOptions defaults() {
            return new QrOptions(ErrorCorrection.M, 2, AUTO_SCALE, 300, BACKGROUND_WHITE);
        }

        /**
         * Parse an alt-text query string such as {@code ec=H&margin=4&bg=transparent}.
         *
         * <p>Unknown options and malformed values are ignored with a warning — parsing never
         * fails, so a typo in a template cannot break document generation.</p>
         *
         * @param query the query string without the leading {@code ?} (may be {@code null})
         * @return parsed options, always non-null
         */
        public static QrOptions parse(String query) {
            if (query == null || query.isBlank()) {
                return defaults();
            }

            QrOptions options = defaults();
            for (String token : query.split("&")) {
                String trimmed = token.trim();
                if (trimmed.isEmpty()) {
                    continue;
                }

                int eq = trimmed.indexOf('=');
                if (eq <= 0) {
                    log.warn("Ignoring malformed QR option '{}' (expected name=value)", trimmed);
                    continue;
                }

                String name = trimmed.substring(0, eq).trim().toLowerCase(Locale.ROOT);
                String value = trimmed.substring(eq + 1).trim();

                switch (name) {
                    case "ec" -> options = options.withErrorCorrection(parseErrorCorrection(value));
                    case "margin" -> options = options.withMargin(parseInt(value, name, MIN_MARGIN, MAX_MARGIN));
                    case "scale" -> options = options.withScale(parseInt(value, name, 1, MAX_PIXELS));
                    case "minpx" -> options = options.withMinPixels(parseInt(value, name, MIN_PIXELS_FLOOR, MAX_PIXELS));
                    case "bg" -> options = options.withBackground(parseBackground(value));
                    default -> log.warn("Ignoring unknown QR option '{}'", name);
                }
            }
            return options;
        }

        private static ErrorCorrection parseErrorCorrection(String value) {
            try {
                return ErrorCorrection.valueOf(value.toUpperCase(Locale.ROOT));
            } catch (IllegalArgumentException e) {
                log.warn("Ignoring invalid QR error correction '{}' (expected L, M, Q or H)", value);
                return null;
            }
        }

        private static int parseInt(String value, String name, int min, int max) {
            try {
                int parsed = Integer.parseInt(value);
                if (parsed < min || parsed > max) {
                    log.warn("QR option '{}={}' is out of range {}-{}, using default", name, value, min, max);
                    return INVALID;
                }
                return parsed;
            } catch (NumberFormatException e) {
                log.warn("Ignoring invalid QR option '{}={}' (expected an integer)", name, value);
                return INVALID;
            }
        }

        private static String parseBackground(String value) {
            String normalized = value.toLowerCase(Locale.ROOT);
            if (BACKGROUND_WHITE.equals(normalized) || BACKGROUND_TRANSPARENT.equals(normalized)) {
                return normalized;
            }
            if (value.matches("#?[0-9a-fA-F]{6}")) {
                return value.startsWith("#") ? value : "#" + value;
            }
            log.warn("Ignoring invalid QR background '{}' (expected white, transparent or #RRGGBB)", value);
            return null;
        }

        /**
         * @param level error correction level
         * @return a copy with the given error correction level (no-op for {@code null})
         */
        public QrOptions withErrorCorrection(ErrorCorrection level) {
            return level == null ? this : new QrOptions(level, marginModules, scale, minPixels, background);
        }

        /**
         * @param modules quiet zone in modules
         * @return a copy with the given quiet zone; out-of-range values are ignored
         */
        public QrOptions withMargin(int modules) {
            return modules == INVALID ? this : new QrOptions(errorCorrection, modules, scale, minPixels, background);
        }

        /**
         * @param pixelsPerModule module size in pixels
         * @return a copy with the given module scale; out-of-range values are ignored
         */
        public QrOptions withScale(int pixelsPerModule) {
            return pixelsPerModule == INVALID
                    ? this
                    : new QrOptions(errorCorrection, marginModules, pixelsPerModule, minPixels, background);
        }

        /**
         * @param pixels minimum edge length in pixels
         * @return a copy with the given minimum edge length; out-of-range values are ignored
         */
        public QrOptions withMinPixels(int pixels) {
            return pixels == INVALID
                    ? this
                    : new QrOptions(errorCorrection, marginModules, scale, pixels, background);
        }

        /**
         * @param backgroundValue {@code white}, {@code transparent} or {@code #RRGGBB}
         * @return a copy with the given background colour; {@code null} keeps the current one
         */
        public QrOptions withBackground(String backgroundValue) {
            return backgroundValue == null
                    ? this
                    : new QrOptions(errorCorrection, marginModules, scale, minPixels, backgroundValue);
        }

        boolean isTransparentBackground() {
            return BACKGROUND_TRANSPARENT.equalsIgnoreCase(background);
        }
    }

    /**
     * Check whether a resolved value asks for QR rendering.
     *
     * @param value the resolved context value (may be {@code null})
     * @return {@code true} if the value starts with {@code qr:}
     */
    public static boolean isQrValue(String value) {
        return value != null && value.startsWith(QR_PREFIX);
    }

    /**
     * Extract the payload from a {@code qr:}-prefixed value.
     *
     * @param value the resolved context value
     * @return the payload after the prefix, trimmed; empty string if there is none
     */
    public static String payloadOf(String value) {
        if (!isQrValue(value)) {
            return "";
        }
        return value.substring(QR_PREFIX.length()).trim();
    }

    /**
     * Encode a payload into a PNG QR code using {@link QrOptions#defaults() default options}.
     *
     * @param payload the text to encode (URL, credentials, plain text)
     * @return PNG bytes, or {@code null} if the payload is blank or encoding failed
     */
    public static byte[] toPng(String payload) {
        return toPng(payload, QrOptions.defaults());
    }

    /**
     * Encode a payload into a PNG QR code with a custom minimum edge length.
     *
     * @param payload   the text to encode
     * @param minPixels target edge length in pixels; the actual size is a whole multiple of
     *                  the symbol size and is never smaller than this value
     * @return PNG bytes, or {@code null} if the payload is blank or encoding failed
     */
    public static byte[] toPng(String payload, int minPixels) {
        return toPng(payload, QrOptions.defaults().withMinPixels(minPixels));
    }

    /**
     * Encode a payload into a PNG QR code.
     *
     * @param payload the text to encode
     * @param options rendering options, or {@code null} for defaults
     * @return PNG bytes, or {@code null} if the payload is blank or encoding failed
     */
    public static byte[] toPng(String payload, QrOptions options) {
        if (payload == null || payload.isBlank()) {
            log.warn("Cannot render QR code: payload is blank");
            return null;
        }
        QrOptions effective = options == null ? QrOptions.defaults() : options;

        try {
            QRCodeWriter writer = new QRCodeWriter();
            Map<EncodeHintType, Object> hints = hints(effective);

            // First pass: minimal symbol (1 px per module) to learn the module count.
            BitMatrix minimal = writer.encode(payload, BarcodeFormat.QR_CODE, 0, 0, hints);
            int modules = minimal.getWidth();
            if (modules <= 0) {
                log.warn("Cannot render QR code: encoder returned an empty symbol for payload '{}'", payload);
                return null;
            }

            int scale = effective.scale() > 0
                    ? effective.scale()
                    : Math.min(40, Math.max(1,
                            (int) Math.ceil((double) Math.max(effective.minPixels(), 1) / modules)));
            int size = modules * scale;

            // Second pass: same payload at the computed pixel size (keeps module edges aligned).
            BitMatrix matrix = writer.encode(payload, BarcodeFormat.QR_CODE, size, size, hints);
            byte[] png = toPngBytes(matrix, effective);
            log.debug("Rendered QR code ({}x{} px, {} modules, scale {}, ec {}, bg {})",
                    matrix.getWidth(), matrix.getHeight(), modules, scale,
                    effective.errorCorrection(), effective.background());
            return png;

        } catch (WriterException | IOException e) {
            log.warn("Failed to render QR code for payload '{}': {}", payload, e.getMessage());
            return null;
        }
    }

    private static Map<EncodeHintType, Object> hints(QrOptions options) {
        Map<EncodeHintType, Object> hints = new EnumMap<>(EncodeHintType.class);
        hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.valueOf(options.errorCorrection().name()));
        hints.put(EncodeHintType.CHARACTER_SET, StandardCharsets.UTF_8.name());
        hints.put(EncodeHintType.MARGIN, options.marginModules());
        return hints;
    }

    private static byte[] toPngBytes(BitMatrix matrix, QrOptions options) throws IOException {
        int width = matrix.getWidth();
        int height = matrix.getHeight();
        boolean transparent = options.isTransparentBackground();

        // Transparent pixels keep white RGB under alpha 0, so pipelines that ignore the alpha
        // channel still see a light background instead of black.
        int backgroundPixel = transparent ? 0x00FFFFFF : backgroundRgb(options.background());
        int imageType = transparent ? BufferedImage.TYPE_INT_ARGB : BufferedImage.TYPE_INT_RGB;

        BufferedImage image = new BufferedImage(width, height, imageType);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                image.setRGB(x, y, matrix.get(x, y) ? 0xFF000000 : backgroundPixel);
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        if (!ImageIO.write(image, "png", out)) {
            throw new IOException("No PNG writer available in this JRE");
        }
        return out.toByteArray();
    }

    private static int backgroundRgb(String background) {
        if (background == null || background.isBlank() || "white".equalsIgnoreCase(background)) {
            return 0xFFFFFF;
        }
        String hex = background.startsWith("#") ? background.substring(1) : background;
        return Integer.parseInt(hex, 16) & 0xFFFFFF;
    }
}
