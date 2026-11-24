package me.muban.docx;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link DocxExporter} — DOCX, PDF, HTML, and TXT export.
 */
@DisplayName("DocxExporter")
class DocxExporterTest {

    @TempDir
    Path tempDir;

    private WordprocessingMLPackage wordPackage;

    @BeforeEach
    void setUp() throws Exception {
        wordPackage = WordprocessingMLPackage.createPackage();
        wordPackage.getMainDocumentPart().addParagraphOfText("Hello, World!");
    }

    // ── DOCX Export ────────────────────────────────────────────────────

    @Nested
    @DisplayName("DOCX export")
    class DocxExport {

        @Test
        @DisplayName("exports a valid .docx file")
        void exportsDocxFile() {
            String outputPath = DocxExporter.exportDocx(wordPackage, tempDir.toString());

            File output = new File(outputPath);
            assertThat(output).exists().isFile();
            assertThat(output.getName()).endsWith(".docx");
            assertThat(output.length()).isGreaterThan(0);
        }

        @Test
        @DisplayName("exported DOCX is loadable")
        void exportedDocxIsLoadable() throws Exception {
            String outputPath = DocxExporter.exportDocx(wordPackage, tempDir.toString());

            WordprocessingMLPackage loaded = WordprocessingMLPackage.load(new File(outputPath));
            assertThat(loaded.getMainDocumentPart().getContent()).isNotEmpty();
        }
    }

    // ── HTML Export ────────────────────────────────────────────────────

    @Nested
    @DisplayName("HTML export")
    class HtmlExport {

        @Test
        @DisplayName("exports a .zip file")
        void exportsZipFile() {
            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString());

            File output = new File(outputPath);
            assertThat(output).exists().isFile();
            assertThat(output.getName()).endsWith(".zip");
            assertThat(output.length()).isGreaterThan(0);
        }

        @Test
        @DisplayName("ZIP contains index.html")
        void zipContainsIndexHtml() throws Exception {
            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString());

            try (ZipFile zip = new ZipFile(outputPath)) {
                boolean hasIndexHtml = zip.stream()
                        .map(ZipEntry::getName)
                        .anyMatch(name -> name.endsWith("/index.html"));
                assertThat(hasIndexHtml)
                        .as("ZIP should contain index.html")
                        .isTrue();
            }
        }

        @Test
        @DisplayName("ZIP contains index.html_files/ directory")
        void zipContainsImageDir() throws Exception {
            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString());

            try (ZipFile zip = new ZipFile(outputPath)) {
                boolean hasImageDir = zip.stream()
                        .map(ZipEntry::getName)
                        .anyMatch(name -> name.contains("index.html_files/"));
                assertThat(hasImageDir)
                        .as("ZIP should contain index.html_files/ directory")
                        .isTrue();
            }
        }

        @Test
        @DisplayName("index.html contains document text")
        void htmlContainsDocumentText() throws Exception {
            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString());

            try (ZipFile zip = new ZipFile(outputPath)) {
                ZipEntry htmlEntry = zip.stream()
                        .filter(e -> e.getName().endsWith("/index.html"))
                        .findFirst()
                        .orElseThrow(() -> new AssertionError("index.html not found in ZIP"));

                try (InputStream is = zip.getInputStream(htmlEntry)) {
                    String html = new String(is.readAllBytes());
                    assertThat(html)
                            .contains("Hello")
                            .containsIgnoringCase("<html");
                }
            }
        }

        @Test
        @DisplayName("temporary HTML directory is cleaned up after export")
        void tempDirectoryCleanedUp() {
            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString());

            // The UUID-named directory used during export should be removed
            File[] remainingDirs = tempDir.toFile().listFiles(File::isDirectory);
            assertThat(remainingDirs)
                    .as("Temporary HTML directory should be cleaned up")
                    .isNullOrEmpty();
        }
    }

    // ── HTML image modes ───────────────────────────────────────────────

    @Nested
    @DisplayName("HTML image modes")
    class HtmlImageModes {

        @Test
        @DisplayName("FILES writes image files and references them by relative path")
        void filesModeWritesImages() throws Exception {
            addTinyImage();

            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString(),
                    new HtmlExportOptions(HtmlImageMode.FILES));

            try (ZipFile zip = new ZipFile(outputPath)) {
                String html = readHtml(zip);
                assertThat(html).contains("src=\"index.html_files/");
                assertThat(imageEntries(zip))
                        .as("ZIP should contain image files")
                        .isNotEmpty();
            }
        }

        @Test
        @DisplayName("DATA_URI inlines images as base64 and writes no image files")
        void dataUriModeInlinesImages() throws Exception {
            addTinyImage();

            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString(),
                    new HtmlExportOptions(HtmlImageMode.DATA_URI));

            try (ZipFile zip = new ZipFile(outputPath)) {
                assertThat(readHtml(zip)).contains("src=\"data:image/png;base64,");
                assertThat(imageEntries(zip))
                        .as("ZIP should not contain image files")
                        .isEmpty();
            }
        }

        @Test
        @DisplayName("CID references images by content ID, ships files and a manifest")
        void cidModeWritesFilesAndManifest() throws Exception {
            addTinyImage();

            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString(),
                    new HtmlExportOptions(HtmlImageMode.CID));

            try (ZipFile zip = new ZipFile(outputPath)) {
                String html = readHtml(zip);
                assertThat(html).contains("src=\"cid:");

                String contentId = html.substring(html.indexOf("src=\"cid:") + 9);
                contentId = contentId.substring(0, contentId.indexOf('"'));
                assertThat(contentId).isNotBlank();

                assertThat(imageEntries(zip)).hasSize(1);

                String manifest = readEntry(zip, HtmlImageManifest.FILE_NAME);
                assertThat(manifest)
                        .contains("\"version\": 1")
                        .contains(contentId)
                        .contains("\"contentType\": \"image/png\"");

                String path = manifest.substring(manifest.indexOf("\"path\": \"") + 9);
                path = path.substring(0, path.indexOf('"'));
                String manifestPath = path;
                assertThat(zip.stream().map(ZipEntry::getName)
                        .filter(name -> name.endsWith("/" + manifestPath))
                        .toList())
                        .as("manifest path should point at an archive entry")
                        .hasSize(1);
            }
        }

        @Test
        @DisplayName("NONE writes no image files")
        void noneModeWritesNoImages() throws Exception {
            addTinyImage();

            String outputPath = DocxExporter.exportHtml(wordPackage, tempDir.toString(),
                    new HtmlExportOptions(HtmlImageMode.NONE));

            try (ZipFile zip = new ZipFile(outputPath)) {
                assertThat(imageEntries(zip))
                        .as("ZIP should not contain image files")
                        .isEmpty();
            }
        }

        private void addTinyImage() throws Exception {
            byte[] png = Base64.getDecoder().decode(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==");

            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(
                    wordPackage, wordPackage.getMainDocumentPart(), png);
            Inline inline = imagePart.createImageInline("pixel.png", "pixel", 1, 1, false);

            P paragraph = new P();
            R run = new R();
            Drawing drawing = new Drawing();
            drawing.getAnchorOrInline().add(inline);
            run.getContent().add(drawing);
            paragraph.getContent().add(run);
            wordPackage.getMainDocumentPart().getContent().add(paragraph);
        }

        private String readHtml(ZipFile zip) throws Exception {
            return readEntry(zip, "index.html");
        }

        private String readEntry(ZipFile zip, String nameSuffix) throws Exception {
            ZipEntry entry = zip.stream()
                    .filter(e -> e.getName().endsWith(nameSuffix))
                    .findFirst()
                    .orElseThrow(() -> new AssertionError(nameSuffix + " not found in ZIP"));
            try (InputStream is = zip.getInputStream(entry)) {
                return new String(is.readAllBytes());
            }
        }

        private List<String> imageEntries(ZipFile zip) {
            return zip.stream()
                    .map(ZipEntry::getName)
                    .filter(name -> name.startsWith("index.html_files/")
                            || name.contains("/index.html_files/"))
                    .filter(name -> name.matches(".*\\.(png|jpe?g|gif|bmp|svg)$"))
                    .toList();
        }
    }

    // ── Unsupported Format ─────────────────────────────────────────────

    @Nested
    @DisplayName("TXT export")
    class TxtExport {

        @Test
        @DisplayName("exports a valid .txt file")
        void exportsTxtFile() {
            String outputPath = DocxExporter.exportTxt(wordPackage, tempDir.toString());

            File output = new File(outputPath);
            assertThat(output).exists().isFile();
            assertThat(output.getName()).endsWith(".txt");
            assertThat(output.length()).isGreaterThan(0);
        }

        @Test
        @DisplayName("exported TXT contains document text")
        void exportedTxtContainsText() throws Exception {
            String outputPath = DocxExporter.exportTxt(wordPackage, tempDir.toString());

            String content = Files.readString(Path.of(outputPath));
            assertThat(content).contains("Hello");
        }

        @Test
        @DisplayName("preserves newlines between paragraphs")
        void preservesNewlinesBetweenParagraphs() throws Exception {
            wordPackage.getMainDocumentPart().addParagraphOfText("Second paragraph");

            String outputPath = DocxExporter.exportTxt(wordPackage, tempDir.toString());

            String content = Files.readString(Path.of(outputPath));
            String[] lines = content.split(System.lineSeparator());
            assertThat(lines).hasSizeGreaterThanOrEqualTo(2);
            assertThat(content).contains("Hello");
            assertThat(content).contains("Second paragraph");
        }

        @Test
        @DisplayName("uses custom line separator from TxtExportOptions")
        void usesCustomLineSeparator() throws Exception {
            wordPackage.getMainDocumentPart().addParagraphOfText("Second paragraph");

            TxtExportOptions opts = new TxtExportOptions(" | ");
            String outputPath = DocxExporter.exportTxt(wordPackage, tempDir.toString(), opts);

            String content = Files.readString(Path.of(outputPath));
            assertThat(content).contains("Hello, World! | Second paragraph");
        }

        @Test
        @DisplayName("soft-wraps long paragraphs when pageWidthInChars is set")
        void softWrapsLongParagraphs() throws Exception {
            // Create a package with one long paragraph
            wordPackage = WordprocessingMLPackage.createPackage();
            wordPackage.getMainDocumentPart().addParagraphOfText(
                    "The quick brown fox jumps over the lazy dog near the river");

            TxtExportOptions opts = new TxtExportOptions(null, false, 30);
            String outputPath = DocxExporter.exportTxt(wordPackage, tempDir.toString(), opts);

            String content = Files.readString(Path.of(outputPath));
            String[] lines = content.split(System.lineSeparator());
            for (String line : lines) {
                // Each line must be <= 30 chars (unless a single word exceeds it)
                assertThat(line.length()).isLessThanOrEqualTo(30);
            }
            // All words must still be present
            assertThat(content).contains("quick").contains("fox").contains("river");
        }
    }

    // ── Unsupported Format ─────────────────────────────────────────────
    class UnsupportedFormat {

        @Test
        @DisplayName("throws UnsupportedOutputFormatException for unknown format")
        void throwsForUnknownFormat() {
            assertThatThrownBy(() ->
                    DocxExporter.exportDocument(wordPackage, "odt", tempDir.toString(), null, null))
                    .isInstanceOf(UnsupportedOutputFormatException.class);
        }
    }
}
