package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link DocxExporter} — DOCX, PDF, and HTML export.
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

    // ── Unsupported Format ─────────────────────────────────────────────

    @Nested
    @DisplayName("Unsupported format")
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
