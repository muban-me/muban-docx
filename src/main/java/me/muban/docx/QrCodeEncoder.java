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
 * <p>Rendered codes use error correction level {@code M} (≈15% recovery), a quiet zone of
 * 2 modules and UTF-8 encoding. Since the picture frame (size and position) is taken from
 * the DOCX template, the PNG is generated with an oversampled module size
 * ({@value #MIN_PIXELS} px minimum) so that scaling down in Word keeps edges crisp.</p>
 *
 * <p>Encoding never throws on bad input: {@code null}, blank payloads and encoder failures
 * are logged and reported as {@code null}, which makes {@link DocxImageReplacer} keep the
 * original template image.</p>
 */
public final class QrCodeEncoder {

    /** Value prefix marking a context value as a QR payload to be rendered. */
    public static final String QR_PREFIX = "qr:";

    private static final Logger log = LoggerFactory.getLogger(QrCodeEncoder.class);

    /** Error correction level — 15% recovery, the usual choice for printed documents. */
    private static final ErrorCorrectionLevel ERROR_CORRECTION = ErrorCorrectionLevel.M;

    /** Quiet zone around the symbol, in modules (QR spec minimum is 4, 2 is safe on white paper). */
    private static final int QUIET_ZONE_MODULES = 2;

    /** Target edge length in pixels — the DOCX frame scales the image down anyway. */
    private static final int MIN_PIXELS = 300;

    /** Upper bound for pixels-per-module, to keep very long payloads from exploding. */
    private static final int MAX_SCALE = 40;

    private QrCodeEncoder() {}

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
     * Encode a payload into a PNG QR code using the default resolution.
     *
     * @param payload the text to encode (URL, credentials, plain text)
     * @return PNG bytes, or {@code null} if the payload is blank or encoding failed
     */
    public static byte[] toPng(String payload) {
        return toPng(payload, MIN_PIXELS);
    }

    /**
     * Encode a payload into a PNG QR code.
     *
     * @param payload   the text to encode
     * @param minPixels target edge length in pixels; the actual size is a whole multiple
     *                  of the symbol size and is never smaller than this value
     * @return PNG bytes, or {@code null} if the payload is blank or encoding failed
     */
    public static byte[] toPng(String payload, int minPixels) {
        if (payload == null || payload.isBlank()) {
            log.warn("Cannot render QR code: payload is blank");
            return null;
        }

        try {
            QRCodeWriter writer = new QRCodeWriter();
            Map<EncodeHintType, Object> hints = hints();

            // First pass: minimal symbol (1 px per module) to learn the module count.
            BitMatrix minimal = writer.encode(payload, BarcodeFormat.QR_CODE, 0, 0, hints);
            int modules = minimal.getWidth();
            if (modules <= 0) {
                log.warn("Cannot render QR code: encoder returned an empty symbol for payload '{}'", payload);
                return null;
            }

            int target = Math.max(minPixels, 1);
            int scale = Math.min(MAX_SCALE, Math.max(1, (int) Math.ceil((double) target / modules)));
            int size = modules * scale;

            // Second pass: same payload at the computed pixel size (keeps module edges aligned).
            BitMatrix matrix = writer.encode(payload, BarcodeFormat.QR_CODE, size, size, hints);
            byte[] png = toPngBytes(matrix);
            log.debug("Rendered QR code for payload '{}' ({}x{} px, {} modules, scale {})",
                    payload, matrix.getWidth(), matrix.getHeight(), modules, scale);
            return png;

        } catch (WriterException | IOException e) {
            log.warn("Failed to render QR code for payload '{}': {}", payload, e.getMessage());
            return null;
        }
    }

    private static Map<EncodeHintType, Object> hints() {
        Map<EncodeHintType, Object> hints = new EnumMap<>(EncodeHintType.class);
        hints.put(EncodeHintType.ERROR_CORRECTION, ERROR_CORRECTION);
        hints.put(EncodeHintType.CHARACTER_SET, StandardCharsets.UTF_8.name());
        hints.put(EncodeHintType.MARGIN, QUIET_ZONE_MODULES);
        return hints;
    }

    private static byte[] toPngBytes(BitMatrix matrix) throws IOException {
        int width = matrix.getWidth();
        int height = matrix.getHeight();
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                image.setRGB(x, y, matrix.get(x, y) ? 0x000000 : 0xFFFFFF);
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        if (!ImageIO.write(image, "png", out)) {
            throw new IOException("No PNG writer available in this JRE");
        }
        return out.toByteArray();
    }
}
