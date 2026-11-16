package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Br;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STVerticalAlignRun;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for {@link DocxRichText} — whitelisted HTML to Word run conversion —
 * and for the {@code ${html:...}} marker handled by {@link MubanDocxEngine}.
 */
@DisplayName("DocxRichText")
class DocxRichTextTest {

    // ==================== helpers ====================

    private static String textOf(R run) {
        return DocxXmlUtils.getRunText(run);
    }

    private static boolean isBreak(R run) {
        return run.getContent().stream()
                .anyMatch(o -> DocxXmlUtils.unwrap(o) instanceof Br);
    }

    private static List<String> texts(List<R> runs) {
        return runs.stream().filter(r -> !isBreak(r)).map(DocxRichTextTest::textOf).toList();
    }

    private static long breaks(List<R> runs) {
        return runs.stream().filter(DocxRichTextTest::isBreak).count();
    }

    private static RPr withFont(String font) {
        RPr rPr = new RPr();
        RFonts fonts = new RFonts();
        fonts.setAscii(font);
        fonts.setHAnsi(font);
        rPr.setRFonts(fonts);
        return rPr;
    }

    // ==================== basics ====================

    @Nested
    @DisplayName("plain text")
    class PlainText {

        @Test
        @DisplayName("null yields no runs")
        void nullValue() {
            assertThat(DocxRichText.toRuns(null, null)).isEmpty();
        }

        @Test
        @DisplayName("empty string yields no runs")
        void emptyValue() {
            assertThat(DocxRichText.toRuns(null, "")).isEmpty();
        }

        @Test
        @DisplayName("text without markup is a single run")
        void noMarkup() {
            List<R> runs = DocxRichText.toRuns(null, "Towarzystwo Ubezpieczeń na Życie");

            assertThat(runs).hasSize(1);
            assertThat(textOf(runs.get(0))).isEqualTo("Towarzystwo Ubezpieczeń na Życie");
        }

        @Test
        @DisplayName("literal markup characters are kept when escaped")
        void escapedMarkup() {
            List<R> runs = DocxRichText.toRuns(null, "&lt;b&gt;nie bold&lt;/b&gt;");

            assertThat(runs).hasSize(1);
            assertThat(textOf(runs.get(0))).isEqualTo("<b>nie bold</b>");
        }
    }

    // ==================== line breaks ====================

    @Nested
    @DisplayName("line breaks")
    class LineBreaks {

        @Test
        @DisplayName("<br/> splits text into two runs with a break between")
        void singleBreak() {
            List<R> runs = DocxRichText.toRuns(null, "ALA<br/>MA");

            assertThat(runs).hasSize(3);
            assertThat(textOf(runs.get(0))).isEqualTo("ALA");
            assertThat(isBreak(runs.get(1))).isTrue();
            assertThat(textOf(runs.get(2))).isEqualTo("MA");
        }

        @Test
        @DisplayName("<br> without slash works the same")
        void breakWithoutSlash() {
            List<R> runs = DocxRichText.toRuns(null, "A<br>B");

            assertThat(breaks(runs)).isEqualTo(1);
            assertThat(texts(runs)).containsExactly("A", "B");
        }

        @Test
        @DisplayName("two consecutive breaks produce an empty line")
        void doubleBreak() {
            List<R> runs = DocxRichText.toRuns(null, "A<br/><br/>B");

            assertThat(breaks(runs)).isEqualTo(2);
        }

        @Test
        @DisplayName("paragraph tags produce a single break")
        void paragraphs() {
            List<R> runs = DocxRichText.toRuns(null, "<p>A</p><p>B</p>");

            assertThat(breaks(runs)).isEqualTo(1);
            assertThat(texts(runs)).containsExactly("A", "B");
        }

        @Test
        @DisplayName("leading and trailing breaks are dropped")
        void edgeBreaks() {
            List<R> runs = DocxRichText.toRuns(null, "<br/>ALA<br/>");

            assertThat(breaks(runs)).isZero();
            assertThat(texts(runs)).containsExactly("ALA");
        }
    }

    // ==================== inline formatting ====================

    @Nested
    @DisplayName("inline formatting")
    class Formatting {

        @Test
        @DisplayName("<b> makes the run bold")
        void bold() {
            List<R> runs = DocxRichText.toRuns(null, "a<b>b</b>c");

            assertThat(texts(runs)).containsExactly("a", "b", "c");
            assertThat(runs.get(0).getRPr()).isNull();
            assertThat(runs.get(1).getRPr().getB()).isNotNull();
            assertThat(runs.get(2).getRPr()).isNull();
        }

        @Test
        @DisplayName("<strong> is an alias of <b>")
        void strong() {
            List<R> runs = DocxRichText.toRuns(null, "<strong>x</strong>");

            assertThat(runs.get(0).getRPr().getB()).isNotNull();
        }

        @Test
        @DisplayName("nested <b><i> produces one run with both attributes")
        void nested() {
            List<R> runs = DocxRichText.toRuns(null, "<b><i>x</i></b>");

            assertThat(runs).hasSize(1);
            assertThat(runs.get(0).getRPr().getB()).isNotNull();
            assertThat(runs.get(0).getRPr().getI()).isNotNull();
        }

        @Test
        @DisplayName("unclosed <b> applies to the remaining text")
        void unclosed() {
            List<R> runs = DocxRichText.toRuns(null, "<b>bold");

            assertThat(runs).hasSize(1);
            assertThat(runs.get(0).getRPr().getB()).isNotNull();
        }

        @Test
        @DisplayName("stray closing tag is ignored")
        void strayClosing() {
            List<R> runs = DocxRichText.toRuns(null, "</b>x");

            assertThat(runs).hasSize(1);
            assertThat(runs.get(0).getRPr()).isNull();
        }

        @Test
        @DisplayName("<u> sets single underline")
        void underline() {
            List<R> runs = DocxRichText.toRuns(null, "<u>x</u>");

            assertThat(runs.get(0).getRPr().getU()).isNotNull();
        }

        @Test
        @DisplayName("<sup> sets superscript")
        void superscript() {
            List<R> runs = DocxRichText.toRuns(null, "BIG<sup>(1)</sup>");

            R marked = runs.get(1);
            assertThat(marked.getRPr().getVertAlign().getVal()).isEqualTo(STVerticalAlignRun.SUPERSCRIPT);
        }

        @Test
        @DisplayName("<sub> sets subscript")
        void subscript() {
            List<R> runs = DocxRichText.toRuns(null, "H<sub>2</sub>O");

            R marked = runs.get(1);
            assertThat(marked.getRPr().getVertAlign().getVal()).isEqualTo(STVerticalAlignRun.SUBSCRIPT);
        }

        @Test
        @DisplayName("base run formatting is preserved on every produced run")
        void baseFormattingPreserved() {
            List<R> runs = DocxRichText.toRuns(withFont("Merriweather Sans Light"), "a<br/>b");

            assertThat(runs).allSatisfy(run ->
                    assertThat(run.getRPr().getRFonts().getAscii()).isEqualTo("Merriweather Sans Light"));
        }
    }

    // ==================== entities ====================

    @Nested
    @DisplayName("entities")
    class Entities {

        @Test
        @DisplayName("named entities are decoded")
        void named() {
            List<R> runs = DocxRichText.toRuns(null, "A&amp;B &lt;C&gt;");

            assertThat(textOf(runs.get(0))).isEqualTo("A&B <C>");
        }

        @Test
        @DisplayName("non-breaking space is preserved")
        void nbsp() {
            List<R> runs = DocxRichText.toRuns(null, "1&nbsp;000");

            assertThat(textOf(runs.get(0))).isEqualTo("1\u00A0000");
        }

        @Test
        @DisplayName("numeric entities are decoded")
        void numeric() {
            List<R> runs = DocxRichText.toRuns(null, "&#261;&#x119;");

            assertThat(textOf(runs.get(0))).isEqualTo("ąę");
        }

        @Test
        @DisplayName("unknown entity stays untouched")
        void unknown() {
            List<R> runs = DocxRichText.toRuns(null, "&foo;");

            assertThat(textOf(runs.get(0))).isEqualTo("&foo;");
        }
    }

    // ==================== robustness ====================

    @Nested
    @DisplayName("robustness")
    class Robustness {

        @Test
        @DisplayName("unknown tags are stripped, text kept")
        void unknownTag() {
            List<R> runs = DocxRichText.toRuns(null, "<span class=\"x\">treść</span>");

            assertThat(texts(runs)).containsExactly("treść");
        }

        @Test
        @DisplayName("unterminated tag is treated as text")
        void unterminatedTag() {
            List<R> runs = DocxRichText.toRuns(null, "a <b c");

            assertThat(textOf(runs.get(0))).isEqualTo("a <b c");
        }

        @Test
        @DisplayName("HTML comment is ignored")
        void comment() {
            List<R> runs = DocxRichText.toRuns(null, "A<!-- ukryte -->B");

            assertThat(texts(runs)).containsExactly("AB");
        }

        @Test
        @DisplayName("over-long value falls back to plain text")
        void tooLong() {
            String payload = "<b>" + "x".repeat(DocxRichText.MAX_LENGTH) + "</b>";

            List<R> runs = DocxRichText.toRuns(null, payload);

            assertThat(runs).hasSize(1);
            assertThat(runs.get(0).getRPr()).isNull();
            assertThat(textOf(runs.get(0))).hasSize(DocxRichText.MAX_LENGTH);
        }

        @Test
        @DisplayName("stripTags removes markup and keeps text")
        void strip() {
            assertThat(DocxRichText.stripTags("<b>A</b><br/>B")).isEqualTo("A B");
            assertThat(DocxRichText.stripTags(null)).isEmpty();
        }
    }

    // ==================== engine integration ====================

    @Nested
    @DisplayName("${html:...} marker in the template")
    class EngineIntegration {

        private P processParagraph(String templateText, Map<String, Object> context) throws Exception {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
            P paragraph = new P();
            R run = new R();
            Text text = new Text();
            text.setValue(templateText);
            text.setSpace("preserve");
            run.getContent().add(text);
            paragraph.getContent().add(run);
            pkg.getMainDocumentPart().getContent().add(paragraph);

            WordprocessingMLPackage processed = MubanDocxEngine.builder()
                    .template(pkg)
                    .rawContext(context)
                    .build()
                    .process();

            return (P) processed.getMainDocumentPart().getContent().stream()
                    .map(DocxXmlUtils::unwrap)
                    .filter(P.class::isInstance)
                    .findFirst()
                    .orElseThrow();
        }

        @Test
        @DisplayName("marked placeholder is rendered as rich text")
        void richPlaceholder() throws Exception {
            P paragraph = processParagraph("${html:footer}", Map.of("footer", "ALA<br/>MA"));

            List<R> runs = paragraph.getContent().stream()
                    .map(DocxXmlUtils::unwrap)
                    .filter(R.class::isInstance)
                    .map(R.class::cast)
                    .toList();

            assertThat(runs).hasSize(3);
            assertThat(textOf(runs.get(0))).isEqualTo("ALA");
            assertThat(isBreak(runs.get(1))).isTrue();
            assertThat(textOf(runs.get(2))).isEqualTo("MA");
        }

        @Test
        @DisplayName("unmarked placeholder keeps markup literal")
        void plainPlaceholder() throws Exception {
            P paragraph = processParagraph("${footer}", Map.of("footer", "ALA<br/>MA"));

            List<R> runs = paragraph.getContent().stream()
                    .map(DocxXmlUtils::unwrap)
                    .filter(R.class::isInstance)
                    .map(R.class::cast)
                    .toList();

            assertThat(runs).hasSize(1);
            assertThat(textOf(runs.get(0))).isEqualTo("ALA<br/>MA");
        }

        @Test
        @DisplayName("text around a marked placeholder is preserved")
        void surroundingText() throws Exception {
            P paragraph = processParagraph("Stopka: ${html:footer} (koniec)",
                    Map.of("footer", "<b>A</b>"));

            List<R> runs = paragraph.getContent().stream()
                    .map(DocxXmlUtils::unwrap)
                    .filter(R.class::isInstance)
                    .map(R.class::cast)
                    .toList();

            assertThat(texts(runs)).containsExactly("Stopka: ", "A", " (koniec)");
        }
    }
}
