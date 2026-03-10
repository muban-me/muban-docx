package me.muban.docx;

import org.docx4j.wml.Br;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link DocxXmlUtils} — low-level OOXML manipulation utilities.
 */
@DisplayName("DocxXmlUtils")
class DocxXmlUtilsTest {

    // ==================== splitRunsAtBreaks() ====================

    @Nested
    @DisplayName("splitRunsAtBreaks() — soft line break normalization")
    class SplitRunsAtBreaks {

        @Test
        @DisplayName("paragraph with no breaks is unchanged")
        void noBreaks_unchanged() {
            P paragraph = new P();
            R run = makeRun("Hello world");
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            assertThat(paragraph.getContent()).hasSize(1);
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(0)))
                    .isEqualTo("Hello world");
        }

        @Test
        @DisplayName("run with text + break + text splits into 3 runs")
        void textBreakText_splitsInto3() {
            P paragraph = new P();
            R run = new R();
            run.getContent().add(makeText("Line one"));
            run.getContent().add(new Br());
            run.getContent().add(makeText("Line two"));
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            assertThat(paragraph.getContent()).hasSize(3);
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(0)))
                    .isEqualTo("Line one");
            assertRunContainsBreak((R) paragraph.getContent().get(1));
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(2)))
                    .isEqualTo("Line two");
        }

        @Test
        @DisplayName("break at start of run produces break-only + text runs")
        void breakAtStart() {
            P paragraph = new P();
            R run = new R();
            run.getContent().add(new Br());
            run.getContent().add(makeText("After break"));
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            assertThat(paragraph.getContent()).hasSize(2);
            assertRunContainsBreak((R) paragraph.getContent().get(0));
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(1)))
                    .isEqualTo("After break");
        }

        @Test
        @DisplayName("break at end of run produces text + break-only runs")
        void breakAtEnd() {
            P paragraph = new P();
            R run = new R();
            run.getContent().add(makeText("Before break"));
            run.getContent().add(new Br());
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            assertThat(paragraph.getContent()).hasSize(2);
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(0)))
                    .isEqualTo("Before break");
            assertRunContainsBreak((R) paragraph.getContent().get(1));
        }

        @Test
        @DisplayName("multiple consecutive breaks each get their own run")
        void consecutiveBreaks() {
            P paragraph = new P();
            R run = new R();
            run.getContent().add(makeText("A"));
            run.getContent().add(new Br());
            run.getContent().add(new Br());
            run.getContent().add(makeText("B"));
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            // Expected: [text "A"], [br], [br], [text "B"]
            assertThat(paragraph.getContent()).hasSize(4);
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(0)))
                    .isEqualTo("A");
            assertRunContainsBreak((R) paragraph.getContent().get(1));
            assertRunContainsBreak((R) paragraph.getContent().get(2));
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(3)))
                    .isEqualTo("B");
        }

        @Test
        @DisplayName("formatting (RPr) is preserved on all split runs")
        void formattingPreserved() {
            P paragraph = new P();
            R run = new R();
            RPr rPr = new RPr();
            RFonts fonts = new RFonts();
            fonts.setAscii("Arial");
            rPr.setRFonts(fonts);
            run.setRPr(rPr);
            run.getContent().add(makeText("A"));
            run.getContent().add(new Br());
            run.getContent().add(makeText("B"));
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            assertThat(paragraph.getContent()).hasSize(3);
            for (Object obj : paragraph.getContent()) {
                R splitRun = (R) obj;
                assertThat(splitRun.getRPr()).isNotNull();
                assertThat(splitRun.getRPr().getRFonts().getAscii()).isEqualTo("Arial");
            }
            // Verify deep copy — not the same object
            R first = (R) paragraph.getContent().get(0);
            R second = (R) paragraph.getContent().get(1);
            assertThat(first.getRPr()).isNotSameAs(second.getRPr());
        }

        @Test
        @DisplayName("runs without breaks in same paragraph are preserved")
        void mixedRunsPreserved() {
            P paragraph = new P();
            R normalRun = makeRun("Normal text");
            R breakRun = new R();
            breakRun.getContent().add(makeText("Line1"));
            breakRun.getContent().add(new Br());
            breakRun.getContent().add(makeText("Line2"));
            R anotherNormal = makeRun("End");

            paragraph.getContent().add(normalRun);
            paragraph.getContent().add(breakRun);
            paragraph.getContent().add(anotherNormal);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            // normalRun(1) + split(3) + anotherNormal(1) = 5
            assertThat(paragraph.getContent()).hasSize(5);
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(0)))
                    .isEqualTo("Normal text");
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(1)))
                    .isEqualTo("Line1");
            assertRunContainsBreak((R) paragraph.getContent().get(2));
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(3)))
                    .isEqualTo("Line2");
            assertThat(DocxXmlUtils.getRunText((R) paragraph.getContent().get(4)))
                    .isEqualTo("End");
        }

        @Test
        @DisplayName("break-only run (no text) produces single break-only run")
        void breakOnlyRun() {
            P paragraph = new P();
            R run = new R();
            run.getContent().add(new Br());
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            assertThat(paragraph.getContent()).hasSize(1);
            assertRunContainsBreak((R) paragraph.getContent().get(0));
        }

        @Test
        @DisplayName("run without RPr splits cleanly (null RPr)")
        void noRPr_splitsCleanly() {
            P paragraph = new P();
            R run = new R();
            // No RPr set
            run.getContent().add(makeText("X"));
            run.getContent().add(new Br());
            run.getContent().add(makeText("Y"));
            paragraph.getContent().add(run);

            DocxXmlUtils.splitRunsAtBreaks(paragraph);

            assertThat(paragraph.getContent()).hasSize(3);
            for (Object obj : paragraph.getContent()) {
                R splitRun = (R) obj;
                assertThat(splitRun.getRPr()).isNull();
            }
        }
    }

    // ==================== getRunText() ====================

    @Nested
    @DisplayName("getRunText()")
    class GetRunText {

        @Test
        @DisplayName("concatenates all Text children")
        void concatenatesText() {
            R run = new R();
            run.getContent().add(makeText("Hello "));
            run.getContent().add(makeText("World"));
            assertThat(DocxXmlUtils.getRunText(run)).isEqualTo("Hello World");
        }

        @Test
        @DisplayName("returns empty string for run with no text")
        void emptyRun() {
            R run = new R();
            assertThat(DocxXmlUtils.getRunText(run)).isEmpty();
        }
    }

    // ==================== setRunText() ====================

    @Nested
    @DisplayName("setRunText()")
    class SetRunText {

        @Test
        @DisplayName("replaces all text with single value")
        void replacesText() {
            R run = new R();
            run.getContent().add(makeText("old1"));
            run.getContent().add(makeText("old2"));

            DocxXmlUtils.setRunText(run, "new value");

            assertThat(DocxXmlUtils.getRunText(run)).isEqualTo("new value");
        }

        @Test
        @DisplayName("null or empty text removes all Text children")
        void nullClearsText() {
            R run = new R();
            run.getContent().add(makeText("something"));

            DocxXmlUtils.setRunText(run, null);

            assertThat(DocxXmlUtils.getRunText(run)).isEmpty();
        }
    }

    // ==================== helpers ====================

    private static R makeRun(String text) {
        R run = new R();
        run.getContent().add(makeText(text));
        return run;
    }

    private static Text makeText(String value) {
        Text t = new Text();
        t.setValue(value);
        return t;
    }

    private static void assertRunContainsBreak(R run) {
        boolean hasBr = run.getContent().stream()
                .map(DocxXmlUtils::unwrap)
                .anyMatch(Br.class::isInstance);
        assertThat(hasBr).as("Run should contain a Br element").isTrue();
    }
}
