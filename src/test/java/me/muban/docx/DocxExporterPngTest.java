package me.muban.docx;

import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests for {@link DocxExporter} PNG export — direct DOCX → XSL-FO → FOP rendering
 * without an intermediate PDF.
 */
@DisplayName("DocxExporter — PNG export")
class DocxExporterPngTest {

    private static final int A4_WIDTH_AT_72_DPI = 595;
    private static final int A4_HEIGHT_AT_72_DPI = 842;
    private static final int DIMENSION_TOLERANCE = 2;

    @TempDir
    Path tempDir;

    private WordprocessingMLPackage wordPackage;

    @BeforeAll
    static void registerTestFont() throws Exception {
        URL font = DocxExporterPngTest.class.getResource("/fonts/DejaVuSans.ttf");
        assertThat(font).as("DejaVu Sans test font").isNotNull();
        PhysicalFonts.addPhysicalFonts("DejaVu Sans", font.toURI());
    }

    @BeforeEach
    void setUp() throws Exception {
        wordPackage = WordprocessingMLPackage.createPackage();
        wordPackage.getMainDocumentPart().addParagraphOfText("Hello, PNG!");
    }

    @Nested
    @DisplayName("single page")
    class SinglePage {

        @Test
        @DisplayName("exports a ZIP with page_1.png at native dimensions")
        void exportsSinglePageZip() throws Exception {
            String zipPath = DocxExporter.exportPng(wordPackage, tempDir.toString());

            assertThat(zipPath).endsWith(".zip");
            assertThat(Path.of(zipPath)).exists().isNotEmptyFile();
            assertThat(pngEntries(zipPath)).containsExactly("page_1.png");

            BufferedImage image = readPage(zipPath, "page_1.png");
            assertThat(image.getWidth()).isBetween(
                    A4_WIDTH_AT_72_DPI - DIMENSION_TOLERANCE,
                    A4_WIDTH_AT_72_DPI + DIMENSION_TOLERANCE);
            assertThat(image.getHeight()).isBetween(
                    A4_HEIGHT_AT_72_DPI - DIMENSION_TOLERANCE,
                    A4_HEIGHT_AT_72_DPI + DIMENSION_TOLERANCE);
        }

        @Test
        @DisplayName("cleans up the temporary PNG directory")
        void cleansUpTemporaryDirectory() throws Exception {
            DocxExporter.exportPng(wordPackage, tempDir.toString());

            try (var entries = Files.list(tempDir)) {
                assertThat(entries.filter(Files::isDirectory).toList()).isEmpty();
            }
        }
    }

    @Nested
    @DisplayName("zoom ratio")
    class ZoomRatio {

        @Test
        @DisplayName("scales output dimensions up and down")
        void scalesOutputDimensions() throws Exception {
            BufferedImage nativeImage = readPage(
                    DocxExporter.exportPng(wordPackage, tempDir.toString()), "page_1.png");
            BufferedImage doubleImage = readPage(
                    DocxExporter.exportPng(wordPackage, tempDir.toString(),
                            new PngExportOptions(2.0f)), "page_1.png");
            BufferedImage halfImage = readPage(
                    DocxExporter.exportPng(wordPackage, tempDir.toString(),
                            new PngExportOptions(0.5f)), "page_1.png");

            assertThat(doubleImage.getWidth()).isBetween(
                    nativeImage.getWidth() * 2 - DIMENSION_TOLERANCE,
                    nativeImage.getWidth() * 2 + DIMENSION_TOLERANCE);
            assertThat(doubleImage.getHeight()).isBetween(
                    nativeImage.getHeight() * 2 - DIMENSION_TOLERANCE,
                    nativeImage.getHeight() * 2 + DIMENSION_TOLERANCE);
            assertThat(halfImage.getWidth()).isBetween(
                    nativeImage.getWidth() / 2 - DIMENSION_TOLERANCE,
                    nativeImage.getWidth() / 2 + DIMENSION_TOLERANCE);
            assertThat(halfImage.getHeight()).isBetween(
                    nativeImage.getHeight() / 2 - DIMENSION_TOLERANCE,
                    nativeImage.getHeight() / 2 + DIMENSION_TOLERANCE);
        }

        @Test
        @DisplayName("rejects non-positive and non-finite ratios")
        void rejectsInvalidRatios() {
            assertThatThrownBy(() -> new PngExportOptions(0f))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("zoomRatio");
            assertThatThrownBy(() -> new PngExportOptions(Float.NaN))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("zoomRatio");
            assertThatThrownBy(() -> new PngExportOptions(Float.POSITIVE_INFINITY))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("zoomRatio");
        }
    }

    @Nested
    @DisplayName("multi-page")
    class MultiPage {

        @Test
        @DisplayName("exports every page of the document")
        void exportsAllPages() throws Exception {
            for (int i = 0; i < 120; i++) {
                wordPackage.getMainDocumentPart().addParagraphOfText("Paragraph " + (i + 1));
            }

            String zipPath = DocxExporter.exportPng(wordPackage, tempDir.toString());
            List<String> pages = pngEntries(zipPath);

            assertThat(pages).hasSizeGreaterThanOrEqualTo(2);
            assertThat(pages.get(0)).isEqualTo("page_1.png");
            assertThat(pages.get(1)).isEqualTo("page_2.png");
            for (String page : pages) {
                assertThat(readPage(zipPath, page).getWidth()).isPositive();
            }
        }
    }

    @Nested
    @DisplayName("engine builder")
    class EngineBuilder {

        @Test
        @DisplayName("passes PNG options through MubanDocxEngine")
        void passesPngOptionsThroughEngine() throws Exception {
            String zipPath = MubanDocxEngine.builder()
                    .template(wordPackage)
                    .outputDir(tempDir.toString())
                    .outputFormat("png")
                    .pngOptions(new PngExportOptions(2.0f))
                    .build()
                    .generate();

            BufferedImage image = readPage(zipPath, "page_1.png");
            assertThat(image.getWidth()).isBetween(
                    A4_WIDTH_AT_72_DPI * 2 - DIMENSION_TOLERANCE,
                    A4_WIDTH_AT_72_DPI * 2 + DIMENSION_TOLERANCE);
            assertThat(image.getHeight()).isBetween(
                    A4_HEIGHT_AT_72_DPI * 2 - DIMENSION_TOLERANCE,
                    A4_HEIGHT_AT_72_DPI * 2 + DIMENSION_TOLERANCE);
        }
    }

    private static BufferedImage readPage(String zipPath, String pageName) throws IOException {
        try (ZipFile zip = new ZipFile(zipPath)) {
            ZipEntry entry = zip.stream()
                    .filter(e -> e.getName().endsWith("/" + pageName))
                    .findFirst()
                    .orElseThrow(() -> new AssertionError(pageName + " not found in " + zipPath));
            try (InputStream is = zip.getInputStream(entry)) {
                BufferedImage image = ImageIO.read(is);
                assertThat(image).as("decoded " + pageName).isNotNull();
                return image;
            }
        }
    }

    private static List<String> pngEntries(String zipPath) throws IOException {
        try (ZipFile zip = new ZipFile(zipPath)) {
            return zip.stream()
                    .map(ZipEntry::getName)
                    .filter(name -> name.endsWith(".png"))
                    .map(name -> name.substring(name.lastIndexOf('/') + 1))
                    .sorted()
                    .toList();
        }
    }
}
