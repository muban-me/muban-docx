package me.muban.docx;

import com.sun.net.httpserver.HttpServer;
import org.docx4j.dml.CTBlip;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.picture.Pic;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.net.InetSocketAddress;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link DocxImageReplacer} — dynamic image replacement in DOCX templates.
 */
@DisplayName("DocxImageReplacer")
class DocxImageReplacerTest {

    // Minimal 1x1 white PNG (67 bytes)
    private static final byte[] TINY_PNG = Base64.getDecoder().decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg=="
    );

    // Minimal 1x1 red PNG (different from TINY_PNG)
    private static final byte[] REPLACEMENT_PNG = Base64.getDecoder().decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg=="
    );

    private static final String REPLACEMENT_DATA_URI =
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg==";

    private static final String TINY_DATA_URI =
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    @TempDir
    Path tempDir;

    private File cacheDir;

    @BeforeEach
    void setUp() {
        cacheDir = tempDir.toFile();
    }

    // ==================== HELPER METHODS ====================

    private WordprocessingMLPackage createDocxWithImage(String altText) throws Exception {
        WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
        MainDocumentPart main = pkg.getMainDocumentPart();

        BinaryPartAbstractImage imgPart =
                BinaryPartAbstractImage.createImagePart(pkg, main, TINY_PNG);

        Inline inline = imgPart.createImageInline("placeholder.png", "placeholder", 1, 1, false);
        inline.getDocPr().setDescr(altText);

        P paragraph = new P();
        R run = new R();
        Drawing drawing = new Drawing();
        drawing.getAnchorOrInline().add(inline);
        run.getContent().add(drawing);
        paragraph.getContent().add(run);
        main.getContent().add(paragraph);

        return pkg;
    }

    private WordprocessingMLPackage createDocxWithAnchoredImage(String altText) throws Exception {
        WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
        MainDocumentPart main = pkg.getMainDocumentPart();

        BinaryPartAbstractImage imgPart =
                BinaryPartAbstractImage.createImagePart(pkg, main, TINY_PNG);

        Inline inline = imgPart.createImageInline("placeholder.png", altText, 1, 1, false);

        Anchor anchor = new Anchor();
        anchor.setDocPr(inline.getDocPr());
        anchor.setGraphic(inline.getGraphic());
        anchor.setExtent(inline.getExtent());
        anchor.setBehindDoc(false);
        anchor.setRelativeHeight(0);

        P paragraph = new P();
        R run = new R();
        Drawing drawing = new Drawing();
        drawing.getAnchorOrInline().add(anchor);
        run.getContent().add(drawing);
        paragraph.getContent().add(run);
        main.getContent().add(paragraph);

        return pkg;
    }

    private byte[] getImageBytes(WordprocessingMLPackage pkg) throws Exception {
        MainDocumentPart main = pkg.getMainDocumentPart();
        for (Object obj : main.getContent()) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (unwrapped instanceof P paragraph) {
                for (Object pObj : paragraph.getContent()) {
                    Object pUnwrapped = DocxXmlUtils.unwrap(pObj);
                    if (pUnwrapped instanceof R run) {
                        for (Object rObj : run.getContent()) {
                            Object rUnwrapped = DocxXmlUtils.unwrap(rObj);
                            if (rUnwrapped instanceof Drawing drawing) {
                                for (Object ai : drawing.getAnchorOrInline()) {
                                    CTBlip blip = null;
                                    if (ai instanceof Inline inline) {
                                        blip = extractBlipFromGraphic(inline.getGraphic());
                                    } else if (ai instanceof Anchor anchor) {
                                        blip = extractBlipFromGraphic(anchor.getGraphic());
                                    }
                                    if (blip != null && blip.getEmbed() != null) {
                                        var rel = main.getRelationshipsPart()
                                                .getRelationshipByID(blip.getEmbed());
                                        var part = main.getRelationshipsPart().getPart(rel);
                                        if (part instanceof org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart bp) {
                                            var buf = bp.getBuffer();
                                            byte[] bytes = new byte[buf.remaining()];
                                            buf.get(bytes);
                                            buf.rewind();
                                            return bytes;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        throw new IllegalStateException("No image found in document");
    }

    private CTBlip extractBlipFromGraphic(Graphic graphic) {
        if (graphic == null || graphic.getGraphicData() == null) return null;
        for (Object any : graphic.getGraphicData().getAny()) {
            Object unwrapped = DocxXmlUtils.unwrap(any);
            if (unwrapped instanceof Pic pic && pic.getBlipFill() != null
                    && pic.getBlipFill().getBlip() != null) {
                return pic.getBlipFill().getBlip();
            }
        }
        return null;
    }

    // ==================== INLINE BASE64 IMAGE REPLACEMENT ====================

    @Nested
    @DisplayName("Inline Base64 image replacement")
    class InlineBase64Tests {

        @Test
        @DisplayName("should replace image when context contains Base64 data URI")
        void shouldReplaceWithBase64DataUri() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:facsimile");

            Map<String, Object> rawContext = Map.of("facsimile", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should replace image from nested images map in context")
        void shouldReplaceFromNestedImagesMap() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:stamp");

            Map<String, Object> rawContext = Map.of(
                    "images", Map.of("stamp", REPLACEMENT_DATA_URI)
            );

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("direct context value takes priority over nested images map")
        void directValueTakesPriority() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:sig");

            Map<String, Object> rawContext = new HashMap<>();
            rawContext.put("sig", REPLACEMENT_DATA_URI);
            rawContext.put("images", Map.of("sig", "data:image/png;base64,WRONG"));

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should handle JPEG data URI")
        void shouldHandleJpegDataUri() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:photo");

            String jpegDataUri = "data:image/jpeg;base64," +
                    Base64.getEncoder().encodeToString(new byte[]{(byte) 0xFF, (byte) 0xD8, (byte) 0xFF});

            Map<String, Object> rawContext = Map.of("photo", jpegDataUri);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should handle data URI with whitespace in Base64 payload")
        void shouldHandleWhitespaceInBase64() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:sig");

            String base64 = Base64.getEncoder().encodeToString(REPLACEMENT_PNG);
            String withNewlines = base64.substring(0, 20) + "\n" + base64.substring(20);
            String dataUri = "data:image/png;base64," + withNewlines;

            Map<String, Object> rawContext = Map.of("sig", dataUri);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }
    }

    // ==================== RELATIVE PATH IMAGE REPLACEMENT ====================

    @Nested
    @DisplayName("Relative path image replacement")
    class RelativePathTests {

        @Test
        @DisplayName("should replace image from relative path in template package")
        void shouldReplaceFromRelativePath() throws Exception {
            Path sigDir = tempDir.resolve("signatures");
            Files.createDirectories(sigDir);
            Files.write(sigDir.resolve("director.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:facsimile");

            Map<String, Object> rawContext = Map.of("facsimile", "signatures/director.png");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should return 0 when relative path file does not exist")
        void shouldNotReplaceWhenFileNotFound() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:facsimile");

            Map<String, Object> rawContext = Map.of("facsimile", "signatures/missing.png");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(0);
            assertThat(getImageBytes(pkg)).isEqualTo(TINY_PNG);
        }

        @Test
        @DisplayName("should block directory traversal in relative path")
        void shouldBlockDirectoryTraversal() throws Exception {
            Path outsideFile = tempDir.getParent().resolve("secret.png");
            Files.write(outsideFile, REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:facsimile");

            Map<String, Object> rawContext = Map.of("facsimile", "../secret.png");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(0);
            assertThat(getImageBytes(pkg)).isEqualTo(TINY_PNG);

            Files.deleteIfExists(outsideFile);
        }
    }

    // ==================== TEMPLATE PATH FALLBACK ====================

    @Nested
    @DisplayName("Template path fallback (key used as relative path)")
    class TemplatePathFallbackTests {

        @Test
        @DisplayName("should load image from key path when no context value")
        void shouldFallbackToKeyAsPath() throws Exception {
            Path assetsDir = tempDir.resolve("assets");
            Files.createDirectories(assetsDir);
            Files.write(assetsDir.resolve("facsimile.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:assets/facsimile.png");

            Map<String, Object> rawContext = Map.of();

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should find image with any directory structure")
        void shouldFindImageInAnyDirectory() throws Exception {
            Path sigDir = tempDir.resolve("signatures");
            Files.createDirectories(sigDir);
            Files.write(sigDir.resolve("stamp.jpg"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:signatures/stamp.jpg");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext(), cacheDir);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should find image in nested directory")
        void shouldFindImageInNestedDirectory() throws Exception {
            Path nestedDir = tempDir.resolve("images").resolve("logos");
            Files.createDirectories(nestedDir);
            Files.write(nestedDir.resolve("company.jpeg"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:images/logos/company.jpeg");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext(), cacheDir);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("context value takes priority over key path")
        void contextTakesPriorityOverKeyPath() throws Exception {
            Path assetsDir = tempDir.resolve("assets");
            Files.createDirectories(assetsDir);
            Files.write(assetsDir.resolve("facsimile.png"), TINY_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:assets/facsimile.png");

            Map<String, Object> rawContext = Map.of("assets/facsimile.png", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should keep original image when key path file does not exist")
        void shouldKeepOriginalWhenNoPathFound() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:missing/file.png");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext(), cacheDir);

            assertThat(replaced).isEqualTo(0);
            assertThat(getImageBytes(pkg)).isEqualTo(TINY_PNG);
        }
    }

    // ==================== ALT TEXT CONVENTION ====================

    @Nested
    @DisplayName("Alt text convention")
    class AltTextConventionTests {

        @Test
        @DisplayName("should only match images with 'image:' prefix in alt text")
        void shouldOnlyMatchImagePrefix() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("Company Logo");

            Map<String, Object> rawContext = Map.of("Company Logo", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should ignore images with empty key after prefix")
        void shouldIgnoreEmptyKey() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext(), cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should ignore images with blank key after prefix")
        void shouldIgnoreBlankKey() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:   ");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext(), cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should trim whitespace in image key")
        void shouldTrimKeyWhitespace() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:  facsimile  ");

            Map<String, Object> rawContext = Map.of("facsimile", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should handle null alt text gracefully")
        void shouldHandleNullAltText() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:facsimile");
            MainDocumentPart main = pkg.getMainDocumentPart();
            for (Object obj : main.getContent()) {
                Object unwrapped = DocxXmlUtils.unwrap(obj);
                if (unwrapped instanceof P p) {
                    for (Object pObj : p.getContent()) {
                        Object pUnwrapped = DocxXmlUtils.unwrap(pObj);
                        if (pUnwrapped instanceof R run) {
                            for (Object rObj : run.getContent()) {
                                Object rUnwrapped = DocxXmlUtils.unwrap(rObj);
                                if (rUnwrapped instanceof Drawing drawing) {
                                    for (Object ai : drawing.getAnchorOrInline()) {
                                        if (ai instanceof Inline inline) {
                                            inline.getDocPr().setDescr(null);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            int replaced = DocxImageReplacer.replaceImages(pkg, Map.of("facsimile", REPLACEMENT_DATA_URI), cacheDir);

            assertThat(replaced).isEqualTo(0);
        }
    }

    // ==================== ANCHORED (FLOATING) IMAGES ====================

    @Nested
    @DisplayName("Anchored image support")
    class AnchoredImageTests {

        @Test
        @DisplayName("should replace anchored (floating) images")
        void shouldReplaceAnchoredImage() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithAnchoredImage("image:facsimile");

            Map<String, Object> rawContext = Map.of("facsimile", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
        }
    }

    // ==================== MULTIPLE IMAGES ====================

    @Nested
    @DisplayName("Multiple images")
    class MultipleImagesTests {

        @Test
        @DisplayName("should replace multiple different placeholder images")
        void shouldReplaceMultipleImages() throws Exception {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
            MainDocumentPart main = pkg.getMainDocumentPart();

            for (String key : List.of("facsimile", "stamp")) {
                BinaryPartAbstractImage imgPart =
                        BinaryPartAbstractImage.createImagePart(pkg, main, TINY_PNG);
                Inline inline = imgPart.createImageInline("img.png", "image:" + key, 1, 1, false);
                P paragraph = new P();
                R run = new R();
                Drawing drawing = new Drawing();
                drawing.getAnchorOrInline().add(inline);
                run.getContent().add(drawing);
                paragraph.getContent().add(run);
                main.getContent().add(paragraph);
            }

            Map<String, Object> rawContext = Map.of(
                    "facsimile", REPLACEMENT_DATA_URI,
                    "stamp", REPLACEMENT_DATA_URI
            );

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(2);
        }

        @Test
        @DisplayName("should skip non-placeholder images while replacing placeholder ones")
        void shouldSkipNonPlaceholderImages() throws Exception {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
            MainDocumentPart main = pkg.getMainDocumentPart();

            for (String altText : List.of("image:facsimile", "Company Logo")) {
                BinaryPartAbstractImage imgPart =
                        BinaryPartAbstractImage.createImagePart(pkg, main, TINY_PNG);
                Inline inline = imgPart.createImageInline("img.png", altText, 1, 1, false);
                P paragraph = new P();
                R run = new R();
                Drawing drawing = new Drawing();
                drawing.getAnchorOrInline().add(inline);
                run.getContent().add(drawing);
                paragraph.getContent().add(run);
                main.getContent().add(paragraph);
            }

            Map<String, Object> rawContext = Map.of("facsimile", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
        }
    }

    // ==================== IMAGES IN TABLE CELLS ====================

    @Nested
    @DisplayName("Images in table cells")
    class TableCellImageTests {

        @Test
        @DisplayName("should replace images inside table cells")
        void shouldReplaceImagesInTableCells() throws Exception {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
            MainDocumentPart main = pkg.getMainDocumentPart();

            BinaryPartAbstractImage imgPart =
                    BinaryPartAbstractImage.createImagePart(pkg, main, TINY_PNG);
            Inline inline = imgPart.createImageInline("img.png", "image:signature", 1, 1, false);

            ObjectFactory factory = new ObjectFactory();
            Tbl table = factory.createTbl();
            Tr row = factory.createTr();
            Tc cell = factory.createTc();

            P paragraph = new P();
            R run = new R();
            Drawing drawing = new Drawing();
            drawing.getAnchorOrInline().add(inline);
            run.getContent().add(drawing);
            paragraph.getContent().add(run);
            cell.getContent().add(paragraph);
            row.getContent().add(cell);
            table.getContent().add(row);
            main.getContent().add(table);

            Map<String, Object> rawContext = Map.of("signature", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
        }
    }

    // ==================== BYTE ARRAY CONTEXT VALUE ====================

    @Nested
    @DisplayName("Byte array context value")
    class ByteArrayTests {

        @Test
        @DisplayName("should accept raw byte[] as context value")
        void shouldAcceptByteArray() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:logo");

            Map<String, Object> rawContext = Map.of("logo", (Object) REPLACEMENT_PNG);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
            assertThat(getImageBytes(pkg)).isEqualTo(REPLACEMENT_PNG);
        }
    }

    // ==================== EDGE CASES ====================

    @Nested
    @DisplayName("Edge cases")
    class EdgeCases {

        @Test
        @DisplayName("should return 0 when document has no images")
        void shouldReturnZeroForNoImages() throws Exception {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
            pkg.getMainDocumentPart().addParagraphOfText("Hello World");

            int replaced = DocxImageReplacer.replaceImages(pkg, Map.of("facsimile", REPLACEMENT_DATA_URI), cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should return 0 with empty context and no assets")
        void shouldReturnZeroWithEmptyContext() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:facsimile");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext(), cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should handle malformed data URI gracefully")
        void shouldHandleMalformedDataUri() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:sig");

            Map<String, Object> rawContext = Map.of("sig", "data:image/png;not-base64,!!invalid!!");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should handle invalid Base64 in data URI gracefully")
        void shouldHandleInvalidBase64() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:sig");

            Map<String, Object> rawContext = Map.of("sig", "data:image/png;base64,NOT_VALID_BASE64!!!");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should not modify non-string, non-byte[] context values")
        void shouldIgnoreNonStringValues() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:sig");

            Map<String, Object> rawContext = Map.of("sig", 12345);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(0);
        }
    }

    // ==================== URL IMAGE FETCHING ====================

    @Nested
    @DisplayName("URL image fetching")
    class UrlImageTests {

        private static HttpServer server;
        private static int port;

        @BeforeAll
        static void startServer() throws Exception {
            server = HttpServer.create(new InetSocketAddress(0), 0);
            port = server.getAddress().getPort();

            server.createContext("/image.png", exchange -> {
                exchange.getResponseHeaders().set("Content-Type", "image/png");
                exchange.sendResponseHeaders(200, TINY_PNG.length);
                exchange.getResponseBody().write(TINY_PNG);
                exchange.getResponseBody().close();
            });

            server.createContext("/photo.jpg", exchange -> {
                exchange.getResponseHeaders().set("Content-Type", "image/jpeg");
                exchange.sendResponseHeaders(200, REPLACEMENT_PNG.length);
                exchange.getResponseBody().write(REPLACEMENT_PNG);
                exchange.getResponseBody().close();
            });

            server.createContext("/not-found", exchange -> {
                exchange.sendResponseHeaders(404, -1);
                exchange.close();
            });

            server.createContext("/error", exchange -> {
                exchange.sendResponseHeaders(500, -1);
                exchange.close();
            });

            server.createContext("/not-image", exchange -> {
                byte[] html = "<html>Not an image</html>".getBytes();
                exchange.getResponseHeaders().set("Content-Type", "text/html");
                exchange.sendResponseHeaders(200, html.length);
                exchange.getResponseBody().write(html);
                exchange.getResponseBody().close();
            });

            server.createContext("/empty", exchange -> {
                exchange.getResponseHeaders().set("Content-Type", "image/png");
                exchange.sendResponseHeaders(200, 0);
                exchange.getResponseBody().close();
            });

            server.createContext("/no-content-type", exchange -> {
                exchange.sendResponseHeaders(200, TINY_PNG.length);
                exchange.getResponseBody().write(TINY_PNG);
                exchange.getResponseBody().close();
            });

            server.start();
        }

        @AfterAll
        static void stopServer() {
            if (server != null) {
                server.stop(0);
            }
            DocxImageReplacer.setHttpClient(null);
        }

        private String url(String path) {
            return "http://localhost:" + port + path;
        }

        @Test
        @DisplayName("should fetch PNG image from URL")
        void shouldFetchPngFromUrl() {
            byte[] result = DocxImageReplacer.loadFromUrl(url("/image.png"));
            assertThat(result).isEqualTo(TINY_PNG);
        }

        @Test
        @DisplayName("should fetch JPEG image from URL")
        void shouldFetchJpegFromUrl() {
            byte[] result = DocxImageReplacer.loadFromUrl(url("/photo.jpg"));
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should return null for HTTP 404")
        void shouldReturnNullForNotFound() {
            byte[] result = DocxImageReplacer.loadFromUrl(url("/not-found"));
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should return null for HTTP 500")
        void shouldReturnNullForServerError() {
            byte[] result = DocxImageReplacer.loadFromUrl(url("/error"));
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should return null for non-image content type")
        void shouldReturnNullForWrongContentType() {
            byte[] result = DocxImageReplacer.loadFromUrl(url("/not-image"));
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should return null for empty body")
        void shouldReturnNullForEmptyBody() {
            byte[] result = DocxImageReplacer.loadFromUrl(url("/empty"));
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should return null for missing content-type header")
        void shouldReturnNullForMissingContentType() {
            byte[] result = DocxImageReplacer.loadFromUrl(url("/no-content-type"));
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should return null for malformed URL")
        void shouldReturnNullForMalformedUrl() {
            byte[] result = DocxImageReplacer.loadFromUrl("http://[invalid-url");
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should return null for unreachable host")
        void shouldReturnNullForUnreachableHost() {
            byte[] result = DocxImageReplacer.loadFromUrl("http://192.0.2.1:1/image.png");
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should replace image when context value is a URL")
        void shouldReplaceImageFromUrl() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:logo");
            Map<String, Object> rawContext = Map.of("logo", url("/image.png"));

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should keep original when URL returns error")
        void shouldKeepOriginalOnUrlError() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:logo");
            Map<String, Object> rawContext = Map.of("logo", url("/not-found"));

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should resolve URL from nested images map")
        void shouldResolveUrlFromNestedImagesMap() {
            Map<String, Object> rawContext = Map.of(
                    "images", Map.of("logo", url("/image.png")));

            byte[] result = DocxImageReplacer.resolveImageBytes("logo", rawContext, cacheDir);
            assertThat(result).isEqualTo(TINY_PNG);
        }
    }

    // ==================== SPEL EXPRESSION SUPPORT ====================

    @Nested
    @DisplayName("SpEL expression in image alt text")
    class SpelImageTests {

        @Test
        @DisplayName("should evaluate simple variable expression to select image")
        void shouldEvaluateSimpleVariable() throws Exception {
            Files.write(tempDir.resolve("photo.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:${photo_path}");
            Map<String, Object> rawContext = Map.of("photo_path", "photo.png");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should evaluate ternary expression for conditional image")
        void shouldEvaluateTernaryExpression() throws Exception {
            Files.createDirectories(tempDir.resolve("assets"));
            Files.write(tempDir.resolve("assets/female.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${gender == 'F' ? 'assets/female.png' : 'assets/male.png'}");
            Map<String, Object> rawContext = Map.of("gender", "F");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should evaluate ternary expression — else branch")
        void shouldEvaluateTernaryElseBranch() throws Exception {
            Files.createDirectories(tempDir.resolve("assets"));
            Files.write(tempDir.resolve("assets/male.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${gender == 'F' ? 'assets/female.png' : 'assets/male.png'}");
            Map<String, Object> rawContext = Map.of("gender", "M");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should evaluate numeric comparison in expression")
        void shouldEvaluateNumericComparison() throws Exception {
            Files.createDirectories(tempDir.resolve("assets"));
            Files.write(tempDir.resolve("assets/exclamation.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${risk > 80 ? 'assets/exclamation.png' : 'assets/info.png'}");
            Map<String, Object> rawContext = Map.of("risk", 95);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should support mixed literal and expression in key")
        void shouldSupportMixedLiteralAndExpression() throws Exception {
            Files.createDirectories(tempDir.resolve("assets/finance"));
            Files.write(tempDir.resolve("assets/finance/stamp.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage("image:assets/${department}/stamp.png");
            Map<String, Object> rawContext = Map.of("department", "finance");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should keep original when expression result file doesn't exist")
        void shouldKeepOriginalWhenExpressionResultNotFound() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${gender == 'F' ? 'assets/female.png' : 'assets/male.png'}");
            Map<String, Object> rawContext = Map.of("gender", "F");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should keep original when expression evaluation fails")
        void shouldKeepOriginalWhenExpressionFails() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${nonexistent_var > 80 ? 'a.png' : 'b.png'}");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext(), cacheDir, null);

            assertThat(replaced).isEqualTo(0);
        }

        @Test
        @DisplayName("should evaluate expression to Base64 data URI")
        void shouldEvaluateExpressionToBase64DataUri() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:${photo}");
            Map<String, Object> rawContext = Map.of("photo", REPLACEMENT_DATA_URI);

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should evaluate expression with equals() method call")
        void shouldEvaluateEqualsMethodCall() throws Exception {
            Files.createDirectories(tempDir.resolve("assets"));
            Files.write(tempDir.resolve("assets/mrs.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${'female'.equals(gender) ? 'assets/mrs.png' : 'assets/mr.png'}");
            Map<String, Object> rawContext = Map.of("gender", "female");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should evaluate expression with double-quoted strings (real-world template)")
        void shouldEvaluateDoubleQuotedStrings() throws Exception {
            Files.createDirectories(tempDir.resolve("assets"));
            Files.write(tempDir.resolve("assets/jaroslaw-niemirowski.png"), REPLACEMENT_PNG);
            Files.write(tempDir.resolve("assets/jaroslaw-niemirow.png"), TINY_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${'niemirowski'.equals(correspondent_surname) ? 'assets/jaroslaw-niemirowski.png' : 'assets/jaroslaw-niemirow.png'}");
            Map<String, Object> rawContext = Map.of("correspondent_surname", "niemirowski");

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should resolve SpEL variable to data URI (unquoted context access)")
        void shouldResolveSpelVariableToDataUri() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${'female'.equals(gender) ? photo_mrs : photo_mr}");
            Map<String, Object> rawContext = Map.of(
                    "gender", "female",
                    "photo_mrs", REPLACEMENT_DATA_URI,
                    "photo_mr", TINY_DATA_URI
            );

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        @Test
        @DisplayName("should resolve SpEL variable to file path (unquoted context access)")
        void shouldResolveSpelVariableToFilePath() throws Exception {
            Files.createDirectories(tempDir.resolve("assets"));
            Files.write(tempDir.resolve("assets/mrs-photo.png"), REPLACEMENT_PNG);

            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${'female'.equals(gender) ? photo_path : other_path}");
            Map<String, Object> rawContext = Map.of(
                    "gender", "female",
                    "photo_path", "assets/mrs-photo.png",
                    "other_path", "assets/other.png"
            );

            int replaced = DocxImageReplacer.replaceImages(pkg, rawContext, cacheDir, null);

            assertThat(replaced).isEqualTo(1);
        }

        // --- Unit tests for evaluateImageExpression ---

        @Test
        @DisplayName("evaluateImageExpression should resolve simple variable")
        void evaluateShouldResolveSimpleVariable() {
            String result = DocxImageReplacer.evaluateImageExpression(
                    "${path}", Map.of("path", "assets/logo.png"), null);
            assertThat(result).isEqualTo("assets/logo.png");
        }

        @Test
        @DisplayName("evaluateImageExpression should resolve ternary")
        void evaluateShouldResolveTernary() {
            String result = DocxImageReplacer.evaluateImageExpression(
                    "${x > 5 ? 'big.png' : 'small.png'}", Map.of("x", 10), null);
            assertThat(result).isEqualTo("big.png");
        }

        @Test
        @DisplayName("evaluateImageExpression should support mixed literal and expression")
        void evaluateShouldSupportMixed() {
            String result = DocxImageReplacer.evaluateImageExpression(
                    "assets/${dept}/stamp.png", Map.of("dept", "hr"), null);
            assertThat(result).isEqualTo("assets/hr/stamp.png");
        }

        @Test
        @DisplayName("evaluateImageExpression should return placeholder on evaluation failure")
        void evaluateShouldReturnPlaceholderOnFailure() {
            String result = DocxImageReplacer.evaluateImageExpression(
                    "${unknown_var.method()}", Map.of(), null);
            assertThat(result).contains("${");
        }

        // --- Unit tests for resolveImageValue ---

        @Test
        @DisplayName("resolveImageValue should decode Base64 data URI")
        void resolveValueShouldDecodeBase64() {
            byte[] result = DocxImageReplacer.resolveImageValue(REPLACEMENT_DATA_URI, cacheDir);
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("resolveImageValue should load from relative path")
        void resolveValueShouldLoadFromPath() throws Exception {
            Files.write(tempDir.resolve("test.png"), REPLACEMENT_PNG);
            byte[] result = DocxImageReplacer.resolveImageValue("test.png", cacheDir);
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("resolveImageValue should return null for non-existent path")
        void resolveValueShouldReturnNullForMissing() {
            byte[] result = DocxImageReplacer.resolveImageValue("no-such-file.png", cacheDir);
            assertThat(result).isNull();
        }
    }

    // ==================== IMAGE KEY EXTRACTION ====================

    @Nested
    @DisplayName("extractImageKeys — template scanning")
    class ExtractImageKeysTests {

        @Test
        @DisplayName("should extract direct image key")
        void shouldExtractDirectKey() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:facsimile");

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).containsExactly("facsimile");
        }

        @Test
        @DisplayName("should extract simple SpEL variable key")
        void shouldExtractSimpleSpelVariable() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:${photo}");

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).containsExactly("photo");
        }

        @Test
        @DisplayName("should extract dot-notation SpEL variable (array field)")
        void shouldExtractDotNotationVariable() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:${items.photo}");

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).containsExactly("items.photo");
        }

        @Test
        @DisplayName("should skip complex SpEL expression")
        void shouldSkipComplexExpression() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage(
                    "image:${gender == 'F' ? 'assets/female.png' : 'assets/male.png'}");

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).isEmpty();
        }

        @Test
        @DisplayName("should skip mixed literal and expression")
        void shouldSkipMixedContent() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("image:assets/${department}/stamp.png");

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).isEmpty();
        }

        @Test
        @DisplayName("should skip images without image: prefix")
        void shouldSkipNonPrefixedImages() throws Exception {
            WordprocessingMLPackage pkg = createDocxWithImage("Company Logo");

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).isEmpty();
        }

        @Test
        @DisplayName("should deduplicate identical keys")
        void shouldDeduplicateKeys() throws Exception {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
            MainDocumentPart main = pkg.getMainDocumentPart();

            for (int i = 0; i < 2; i++) {
                BinaryPartAbstractImage imagePart =
                        BinaryPartAbstractImage.createImagePart(pkg, main, TINY_PNG);
                Inline inline = imagePart.createImageInline("placeholder", "placeholder",
                        900 + i, 900 + i, false);
                inline.getDocPr().setDescr("image:stamp");

                R run = new R();
                Drawing drawing = new Drawing();
                drawing.getAnchorOrInline().add(inline);
                run.getContent().add(drawing);

                P para = new P();
                para.getContent().add(run);
                main.getContent().add(para);
            }

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).containsExactly("stamp");
        }

        @Test
        @DisplayName("should extract multiple different keys")
        void shouldExtractMultipleKeys() throws Exception {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
            MainDocumentPart main = pkg.getMainDocumentPart();

            String[] altTexts = {"image:signature", "image:${logo}", "image:stamp"};
            int idCounter = 800;
            for (String altText : altTexts) {
                BinaryPartAbstractImage imagePart =
                        BinaryPartAbstractImage.createImagePart(pkg, main, TINY_PNG);
                Inline inline = imagePart.createImageInline("placeholder", "placeholder",
                        idCounter, idCounter + 1, false);
                inline.getDocPr().setDescr(altText);

                R run = new R();
                Drawing drawing = new Drawing();
                drawing.getAnchorOrInline().add(inline);
                run.getContent().add(drawing);

                P para = new P();
                para.getContent().add(run);
                main.getContent().add(para);
                idCounter += 2;
            }

            Set<String> keys = DocxImageReplacer.extractImageKeys(pkg);

            assertThat(keys).containsExactlyInAnyOrder("signature", "logo", "stamp");
        }
    }

    // ==================== SECURITY ====================

    @Nested
    @DisplayName("Path security")
    class PathSecurityTests {

        @Test
        @DisplayName("isPathSafe should accept paths within base directory")
        void safePathWithinBaseDir() {
            Path base = tempDir;
            Path safe = tempDir.resolve("assets/image.png");
            assertThat(DocxImageReplacer.isPathSafe(safe, base)).isTrue();
        }

        @Test
        @DisplayName("isPathSafe should reject paths above base directory")
        void unsafePathAboveBaseDir() {
            Path base = tempDir;
            Path unsafe = tempDir.resolve("../outside.png").normalize();
            assertThat(DocxImageReplacer.isPathSafe(unsafe, base)).isFalse();
        }

        @Test
        @DisplayName("isPathSafe should accept equal paths")
        void safeEqualPath() {
            assertThat(DocxImageReplacer.isPathSafe(tempDir, tempDir)).isTrue();
        }

        @Test
        @DisplayName("isPathSafe should reject path with .. that escapes")
        void unsafePathWithDotDot() {
            Path basePath = tempDir.resolve("cache");
            Path escapePath = tempDir.resolve("cache/../../../etc/passwd");
            assertThat(DocxImageReplacer.isPathSafe(escapePath, basePath)).isFalse();
        }
    }

    // ==================== resolveImageBytes (package-private) ====================

    @Nested
    @DisplayName("resolveImageBytes resolution cascade")
    class ResolveImageBytesTests {

        @Test
        @DisplayName("should return null for completely unknown key")
        void shouldReturnNullForUnknownKey() {
            byte[] result = DocxImageReplacer.resolveImageBytes("unknown", Map.of(), cacheDir);
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should decode Base64 data URI from context")
        void shouldDecodeBase64FromContext() {
            byte[] result = DocxImageReplacer.resolveImageBytes(
                    "sig", Map.of("sig", REPLACEMENT_DATA_URI), cacheDir);
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should read file from relative path")
        void shouldReadFromRelativePath() throws Exception {
            Files.write(tempDir.resolve("test.png"), REPLACEMENT_PNG);

            byte[] result = DocxImageReplacer.resolveImageBytes(
                    "sig", Map.of("sig", "test.png"), cacheDir);
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should fall back to key as relative path")
        void shouldFallbackToKeyAsPath() throws Exception {
            Files.write(tempDir.resolve("sig.png"), REPLACEMENT_PNG);

            byte[] result = DocxImageReplacer.resolveImageBytes("sig.png", Map.of(), cacheDir);
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should return null when key path file does not exist")
        void shouldReturnNullWhenKeyPathNotFound() {
            byte[] result = DocxImageReplacer.resolveImageBytes("missing/file.png", Map.of(), cacheDir);
            assertThat(result).isNull();
        }

        @Test
        @DisplayName("should accept byte[] value directly from context")
        void shouldAcceptByteArrayFromContext() {
            byte[] result = DocxImageReplacer.resolveImageBytes(
                    "sig", Map.of("sig", (Object) REPLACEMENT_PNG), cacheDir);
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }

        @Test
        @DisplayName("should prefer Base64 data URI over URL")
        void shouldPreferBase64OverUrl() {
            byte[] result = DocxImageReplacer.resolveImageBytes(
                    "sig", Map.of("sig", REPLACEMENT_DATA_URI), cacheDir);
            assertThat(result).isEqualTo(REPLACEMENT_PNG);
        }
    }

    // ==================== UTILITY ====================

    private static Map<String, Object> rawContext() {
        return Map.of();
    }
}
