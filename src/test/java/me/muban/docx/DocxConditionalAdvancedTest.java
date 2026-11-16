package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.*;

import static me.muban.docx.DocxTestSupport.createDocx;
import static me.muban.docx.DocxTestSupport.createParagraph;
import static me.muban.docx.DocxTestSupport.extractTexts;
import static org.assertj.core.api.Assertions.*;

/**
 * Tests for the structural side of {@link DocxConditionalProcessor}: nested blocks and
 * {@code #{elseif}} branches.
 */
@DisplayName("DocxConditionalProcessor — nesting and branches")
class DocxConditionalAdvancedTest {

    // ==================== NESTED BLOCKS ====================

    @Nested
    @DisplayName("Nested conditional blocks")
    class NestedConditionalTests {

        @Test
        @DisplayName("Both outer and inner condition true")
        void bothTrue() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if outer}",
                    "Outer start",
                    "#{if inner}",
                    "Inner content",
                    "#{fi}",
                    "Outer end",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("outer", true, "inner", true), null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly(
                    "Outer start", "Inner content", "Outer end");
        }

        @Test
        @DisplayName("Outer true, inner false keeps inner else branch")
        void outerTrueInnerFalseKeepsElse() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if outer}",
                    "Outer start",
                    "#{if inner}",
                    "Inner yes",
                    "#{else}",
                    "Inner no",
                    "#{fi}",
                    "Outer end",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("outer", true, "inner", false), null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly(
                    "Outer start", "Inner no", "Outer end");
        }

        @Test
        @DisplayName("Outer false removes the whole block including nested markers")
        void outerFalseRemovesEverything() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if outer}",
                    "#{if inner}",
                    "Inner content",
                    "#{fi}",
                    "#{fi}",
                    "After"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("outer", false), null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Before", "After");
        }

        @Test
        @DisplayName("Nested block inside the else branch")
        void nestedInElseBranch() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if a}",
                    "A branch",
                    "#{else}",
                    "#{if b}",
                    "B yes",
                    "#{else}",
                    "B no",
                    "#{fi}",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("a", false, "b", true), null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly("B yes");
        }

        @Test
        @DisplayName("Three levels of nesting")
        void threeLevels() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if level1}",
                    "L1",
                    "#{if level2}",
                    "L2",
                    "#{if level3}",
                    "L3",
                    "#{fi}",
                    "#{fi}",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("level1", true, "level2", true, "level3", true), null);

            assertThat(processed).isEqualTo(3);
            assertThat(extractTexts(content)).containsExactly("L1", "L2", "L3");
        }

        @Test
        @DisplayName("Nested block that stops at the middle level")
        void threeLevelsStoppedInMiddle() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if level1}",
                    "L1",
                    "#{if level2}",
                    "L2",
                    "#{if level3}",
                    "L3",
                    "#{fi}",
                    "#{fi}",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            DocxConditionalProcessor.processConditionals(
                    content, Map.of("level1", true, "level2", false, "level3", true), null);

            assertThat(extractTexts(content)).containsExactly("L1");
        }

        @Test
        @DisplayName("Nested block inside a table cell")
        void nestedInTableCell() throws Exception {
            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainPart = docx.getMainDocumentPart();
            mainPart.getContent().clear();

            Tbl table = new Tbl();
            Tr row = new Tr();
            Tc cell = new Tc();
            cell.getContent().add(createParagraph("#{if a}"));
            cell.getContent().add(createParagraph("#{if b}"));
            cell.getContent().add(createParagraph("Both"));
            cell.getContent().add(createParagraph("#{fi}"));
            cell.getContent().add(createParagraph("#{fi}"));
            row.getContent().add(cell);
            table.getContent().add(row);
            mainPart.getContent().add(table);

            int processed = DocxConditionalProcessor.processConditionals(
                    mainPart.getContent(), Map.of("a", true, "b", true), null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(cell.getContent())).containsExactly("Both");
        }

        @Test
        @DisplayName("Unbalanced inner block leaves the outer marker in place")
        void unbalancedInnerBlock() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if a}",
                    "#{if b}",
                    "Content",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            DocxConditionalProcessor.processConditionals(
                    content, Map.of("a", true, "b", true), null);

            // The inner #{if b} swallows the only #{fi}, so the outer block is reported as
            // unmatched and left untouched; the inner block is then resolved on its own.
            assertThat(extractTexts(content)).containsExactly("#{if a}", "Content");
        }

        @Test
        @DisplayName("Nested condition in a discarded branch is not evaluated")
        void discardedBranchIsNotEvaluated() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if a}",
                    "Kept",
                    "#{else}",
                    "#{if unknown_variable > 1}",
                    "Hidden",
                    "#{fi}",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("a", true), null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Kept");
        }
    }

    // ==================== ELSEIF BRANCHES ====================

    @Nested
    @DisplayName("#{elseif} branches")
    class ElseIfTests {

        private WordprocessingMLPackage switchDocx() throws Exception {
            return createDocx(
                    "#{if type == 'A'}",
                    "Alpha",
                    "#{elseif type == 'B'}",
                    "Beta",
                    "#{elseif type == 'C'}",
                    "Gamma",
                    "#{else}",
                    "Other",
                    "#{fi}"
            );
        }

        @Test
        @DisplayName("First branch wins")
        void firstBranchWins() throws Exception {
            WordprocessingMLPackage docx = switchDocx();
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("type", "A"), null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Alpha");
        }

        @Test
        @DisplayName("Second branch wins")
        void secondBranchWins() throws Exception {
            WordprocessingMLPackage docx = switchDocx();
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, Map.of("type", "B"), null);

            assertThat(extractTexts(content)).containsExactly("Beta");
        }

        @Test
        @DisplayName("Third branch wins")
        void thirdBranchWins() throws Exception {
            WordprocessingMLPackage docx = switchDocx();
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, Map.of("type", "C"), null);

            assertThat(extractTexts(content)).containsExactly("Gamma");
        }

        @Test
        @DisplayName("Falls back to #{else} when nothing matches")
        void fallsBackToElse() throws Exception {
            WordprocessingMLPackage docx = switchDocx();
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, Map.of("type", "Z"), null);

            assertThat(extractTexts(content)).containsExactly("Other");
        }

        @Test
        @DisplayName("Without #{else} nothing survives when nothing matches")
        void noMatchNoElseRemovesBlock() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if a}",
                    "A",
                    "#{elseif b}",
                    "B",
                    "#{fi}",
                    "After"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            DocxConditionalProcessor.processConditionals(
                    content, Map.of("a", false, "b", false), null);

            assertThat(extractTexts(content)).containsExactly("Before", "After");
        }

        @Test
        @DisplayName("#{elif} is accepted as an alias")
        void elifAlias() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if a}",
                    "A",
                    "#{elif b}",
                    "B",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            DocxConditionalProcessor.processConditionals(
                    content, Map.of("a", false, "b", true), null);

            assertThat(extractTexts(content)).containsExactly("B");
        }

        @Test
        @DisplayName("#{elseif} after #{else} is dropped, its body joins the else branch")
        void elseIfAfterElseIsDropped() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if a}",
                    "A",
                    "#{else}",
                    "Else content",
                    "#{elseif b}",
                    "B content",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            DocxConditionalProcessor.processConditionals(
                    content, Map.of("a", false, "b", true), null);

            assertThat(extractTexts(content)).containsExactly("Else content", "B content");
        }

        @Test
        @DisplayName("Nested block inside an #{elseif} branch")
        void nestedInsideElseIf() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if a}",
                    "A",
                    "#{elseif b}",
                    "#{if c}",
                    "B and C",
                    "#{else}",
                    "B only",
                    "#{fi}",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(
                    content, Map.of("a", false, "b", true, "c", false), null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly("B only");
        }

        @Test
        @DisplayName("Variables from #{elseif} expressions are extracted")
        void extractsElseIfVariables() throws Exception {
            WordprocessingMLPackage docx = switchDocx();

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            // 'type' is referenced by the #{if} and both #{elseif} branches (deduplicated)
            assertThat(vars).containsExactly("type");
        }

        @Test
        @DisplayName("ELSEIF_PATTERN matches valid markers")
        void elseIfPatternMatches() {
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("#{elseif b}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("#{elseif type == 'B'}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("#{elif b}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("  #{elseif b}  ").matches()).isTrue();
        }

        @Test
        @DisplayName("ELSEIF_PATTERN rejects invalid markers")
        void elseIfPatternRejects() {
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("#{elseif}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("#{else}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("#{if b}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.ELSEIF_PATTERN.matcher("Some text #{elseif b}").matches()).isFalse();
        }
    }
}
