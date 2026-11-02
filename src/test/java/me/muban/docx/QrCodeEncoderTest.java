package me.muban.docx;

import com.google.zxing.BinaryBitmap;
import com.google.zxing.MultiFormatReader;
import com.google.zxing.Result;
import com.google.zxing.RGBLuminanceSource;
import com.google.zxing.common.HybridBinarizer;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for {@link QrCodeEncoder} — QR code rendering for dynamic image placeholders.
 */
@DisplayName("QrCodeEncoder")
class QrCodeEncoderTest {

    private static final byte[] PNG_MAGIC = {(byte) 0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A};

    private static final String PAYMENT_URL =
            "https://www.warta.pl/obsluga/platnosci?nr=920046666381&kwota=2381.99";

    private static BufferedImage readPng(byte[] png) throws Exception {
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(png));
        assertThat(image).as("PNG should be readable").isNotNull();
        return image;
    }

    private static String decode(byte[] png) throws Exception {
        BufferedImage image = readPng(png);
        int width = image.getWidth();
        int height = image.getHeight();
        int[] pixels = image.getRGB(0, 0, width, height, null, 0, width);

        BinaryBitmap bitmap = new BinaryBitmap(
                new HybridBinarizer(new RGBLuminanceSource(width, height, pixels)));
        Result result = new MultiFormatReader().decode(bitmap);
        return result.getText();
    }

    @Nested
    @DisplayName("Value prefix handling")
    class PrefixHandling {

        @Test
        @DisplayName("should recognise qr: prefix")
        void shouldRecognisePrefix() {
            assertThat(QrCodeEncoder.isQrValue("qr:https://example.com")).isTrue();
            assertThat(QrCodeEncoder.isQrValue("https://example.com")).isFalse();
            assertThat(QrCodeEncoder.isQrValue("data:image/png;base64,AAAA")).isFalse();
            assertThat(QrCodeEncoder.isQrValue(null)).isFalse();
        }

        @Test
        @DisplayName("should extract and trim payload")
        void shouldExtractPayload() {
            assertThat(QrCodeEncoder.payloadOf("qr:  https://example.com  "))
                    .isEqualTo("https://example.com");
            assertThat(QrCodeEncoder.payloadOf("qr:")).isEmpty();
            assertThat(QrCodeEncoder.payloadOf("https://example.com")).isEmpty();
        }
    }

    @Nested
    @DisplayName("PNG rendering")
    class Rendering {

        @Test
        @DisplayName("should render a PNG image")
        void shouldRenderPng() {
            byte[] png = QrCodeEncoder.toPng(PAYMENT_URL);

            assertThat(png).isNotNull();
            assertThat(png.length).isGreaterThan(PNG_MAGIC.length);
            assertThat(png).startsWith(PNG_MAGIC);
        }

        @Test
        @DisplayName("should render a square image of at least the requested size")
        void shouldRenderSquareImage() throws Exception {
            byte[] png = QrCodeEncoder.toPng(PAYMENT_URL, 500);

            BufferedImage image = readPng(png);
            assertThat(image.getWidth()).isEqualTo(image.getHeight());
            assertThat(image.getWidth()).isGreaterThanOrEqualTo(500);
        }

        @Test
        @DisplayName("should round-trip the payload through a QR decoder")
        void shouldRoundTripPayload() throws Exception {
            byte[] png = QrCodeEncoder.toPng(PAYMENT_URL);

            assertThat(decode(png)).isEqualTo(PAYMENT_URL);
        }

        @Test
        @DisplayName("should round-trip non-ASCII payload using UTF-8")
        void shouldRoundTripUtf8() throws Exception {
            String payload = "https://example.com/płatność?opis=zażółć gęślą jaźń";

            byte[] png = QrCodeEncoder.toPng(payload);

            assertThat(decode(png)).isEqualTo(payload);
        }
    }

    @Nested
    @DisplayName("Invalid payloads")
    class InvalidPayloads {

        @Test
        @DisplayName("should return null for null payload")
        void shouldReturnNullForNull() {
            assertThat(QrCodeEncoder.toPng(null)).isNull();
        }

        @Test
        @DisplayName("should return null for blank payload")
        void shouldReturnNullForBlank() {
            assertThat(QrCodeEncoder.toPng("   ")).isNull();
        }
    }
}
