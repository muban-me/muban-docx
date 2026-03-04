package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.*;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link DocxConditionalProcessor} — conditional block processing in DOCX templates.
 */
@DisplayName("DocxConditionalProcessor")
class DocxConditionalProcessorTest {

    // ==================== HELPER METHODS ====================

    /**
     * Create a paragraph with the given text content.
     */
    private static P createParagraph(String text) {
        P paragraph = new P();
        R run = new R();
        Text t = new Text();
        t.setValue(text);
        t.setSpace("preserve");
        run.getContent().add(t);
        paragraph.getContent().add(run);
        return paragraph;
    }

    /**
     * Create a simple DOCX document with the given paragraphs.
     */
    private static WordprocessingMLPackage createDocx(String... paragraphTexts) throws Exception {
        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainPart = wordPackage.getMainDocumentPart();
        mainPart.getContent().clear();
        for (String text : paragraphTexts) {
            mainPart.getContent().add(createParagraph(text));
        }
        return wordPackage;
    }

    /**
     * Extract all text from a content list as a list of paragraph texts.
     */
    private static List<String> extractTexts(List<Object> content) {
        List<String> texts = new ArrayList<>();
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (unwrapped instanceof P paragraph) {
                texts.add(DocxXmlUtils.extractAllText(List.of(paragraph)));
            }
        }
        return texts;
    }

    // ==================== BASIC CONDITIONAL TESTS ====================

    @Nested
    @DisplayName("Basic conditional blocks")
    class BasicConditionalTests {

        @Test
        @DisplayName("True condition keeps content between markers")
        void trueConditionKeepsContent() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show_notice}",
                    "Legal notice content",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show_notice", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly(
                    "Before", "Legal notice content", "After"
            );
        }

        @Test
        @DisplayName("False condition removes content between markers")
        void falseConditionRemovesContent() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show_notice}",
                    "Legal notice content",
                    "Second notice line",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show_notice", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Before", "After");
        }

        @Test
        @DisplayName("Empty document returns zero blocks")
        void emptyDocumentReturnsZero() throws Exception {
            WordprocessingMLPackage docx = createDocx();
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, Map.of(), null);

            assertThat(processed).isEqualTo(0);
        }

        @Test
        @DisplayName("Document without conditionals is unchanged")
        void noConditionalsUnchanged() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "First paragraph",
                    "Second paragraph",
                    "Third paragraph"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(content, Map.of(), null);

            assertThat(processed).isEqualTo(0);
            assertThat(extractTexts(content)).containsExactly(
                    "First paragraph", "Second paragraph", "Third paragraph"
            );
        }
    }

    // ==================== EXPRESSION EVALUATION TESTS ====================

    @Nested
    @DisplayName("Expression evaluation")
    class ExpressionEvaluationTests {

        @Test
        @DisplayName("Numeric comparison: debt > 1000")
        void numericComparison() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Dear customer,",
                    "#{if customer_debt > 1000}",
                    "You have an outstanding debt. Please pay immediately.",
                    "#{fi}",
                    "Kind regards"
            );

            Map<String, Object> context = Map.of("customer_debt", 1500);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly(
                    "Dear customer,",
                    "You have an outstanding debt. Please pay immediately.",
                    "Kind regards"
            );
        }

        @Test
        @DisplayName("Numeric comparison: debt below threshold removes block")
        void numericComparisonBelowThreshold() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Dear customer,",
                    "#{if customer_debt > 1000}",
                    "You have an outstanding debt. Please pay immediately.",
                    "#{fi}",
                    "Kind regards"
            );

            Map<String, Object> context = Map.of("customer_debt", 500);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly(
                    "Dear customer,", "Kind regards"
            );
        }

        @Test
        @DisplayName("String equality check")
        void stringEquality() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if status == 'VIP'}",
                    "Welcome to the VIP lounge!",
                    "#{fi}"
            );

            Map<String, Object> context = Map.of("status", "VIP");
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Welcome to the VIP lounge!");
        }

        @Test
        @DisplayName("Boolean variable directly")
        void booleanVariable() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if is_premium}",
                    "Premium features enabled",
                    "#{fi}"
            );

            Map<String, Object> context = Map.of("is_premium", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Premium features enabled");
        }

        @Test
        @DisplayName("Method call in expression")
        void methodCallExpression() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if items.size() > 0}",
                    "Here are your items:",
                    "#{fi}"
            );

            Map<String, Object> context = Map.of("items", List.of("a", "b"));
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Here are your items:");
        }

        @Test
        @DisplayName("Missing variable evaluates to false")
        void missingVariableEvaluatesToFalse() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if nonexistent_var}",
                    "This should not appear",
                    "#{fi}",
                    "After"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            DocxConditionalProcessor.processConditionals(content, Map.of(), null);

            assertThat(extractTexts(content)).containsExactly("Before", "After");
        }
    }

    // ==================== BOOLEAN COERCION TESTS ====================

    @Nested
    @DisplayName("Boolean coercion")
    class BooleanCoercionTests {

        @Test
        @DisplayName("'true' string coerces to true")
        void trueStringCoercesToTrue() {
            assertThat(DocxConditionalProcessor.coerceToBoolean("true")).isTrue();
            assertThat(DocxConditionalProcessor.coerceToBoolean("TRUE")).isTrue();
            assertThat(DocxConditionalProcessor.coerceToBoolean("True")).isTrue();
        }

        @Test
        @DisplayName("'false' string coerces to false")
        void falseStringCoercesToFalse() {
            assertThat(DocxConditionalProcessor.coerceToBoolean("false")).isFalse();
            assertThat(DocxConditionalProcessor.coerceToBoolean("FALSE")).isFalse();
        }

        @Test
        @DisplayName("Null and empty coerce to false")
        void nullAndEmptyCoerceToFalse() {
            assertThat(DocxConditionalProcessor.coerceToBoolean(null)).isFalse();
            assertThat(DocxConditionalProcessor.coerceToBoolean("")).isFalse();
        }

        @Test
        @DisplayName("Numeric zero coerces to false")
        void numericZeroCoercesToFalse() {
            assertThat(DocxConditionalProcessor.coerceToBoolean("0")).isFalse();
            assertThat(DocxConditionalProcessor.coerceToBoolean("0.0")).isFalse();
        }

        @Test
        @DisplayName("Non-zero numeric coerces to true")
        void nonZeroNumericCoercesToTrue() {
            assertThat(DocxConditionalProcessor.coerceToBoolean("1")).isTrue();
            assertThat(DocxConditionalProcessor.coerceToBoolean("42")).isTrue();
            assertThat(DocxConditionalProcessor.coerceToBoolean("-1")).isTrue();
        }

        @Test
        @DisplayName("Non-empty non-boolean string coerces to true")
        void nonEmptyStringCoercesToTrue() {
            assertThat(DocxConditionalProcessor.coerceToBoolean("hello")).isTrue();
            assertThat(DocxConditionalProcessor.coerceToBoolean("VIP")).isTrue();
        }
    }

    // ==================== MULTIPLE BLOCKS TESTS ====================

    @Nested
    @DisplayName("Multiple conditional blocks")
    class MultipleBlocksTests {

        @Test
        @DisplayName("Two independent blocks, both true")
        void twoBlocksBothTrue() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Header",
                    "#{if show_a}",
                    "Block A",
                    "#{fi}",
                    "Middle",
                    "#{if show_b}",
                    "Block B",
                    "#{fi}",
                    "Footer"
            );

            Map<String, Object> context = Map.of("show_a", true, "show_b", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly(
                    "Header", "Block A", "Middle", "Block B", "Footer"
            );
        }

        @Test
        @DisplayName("Two independent blocks, first true second false")
        void twoBlocksMixed() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Header",
                    "#{if show_a}",
                    "Block A",
                    "#{fi}",
                    "Middle",
                    "#{if show_b}",
                    "Block B",
                    "#{fi}",
                    "Footer"
            );

            Map<String, Object> context = Map.of("show_a", true, "show_b", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly(
                    "Header", "Block A", "Middle", "Footer"
            );
        }

        @Test
        @DisplayName("Two independent blocks, both false")
        void twoBlocksBothFalse() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Header",
                    "#{if show_a}",
                    "Block A content",
                    "#{fi}",
                    "Middle",
                    "#{if show_b}",
                    "Block B content",
                    "#{fi}",
                    "Footer"
            );

            Map<String, Object> context = Map.of("show_a", false, "show_b", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly("Header", "Middle", "Footer");
        }
    }

    // ==================== ELSE BLOCK TESTS ====================

    @Nested
    @DisplayName("If/else/fi blocks")
    class ElseBlockTests {

        @Test
        @DisplayName("True condition keeps if-branch, removes else-branch")
        void trueKeepsIfBranch() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show}",
                    "If content",
                    "#{else}",
                    "Else content",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Before", "If content", "After");
        }

        @Test
        @DisplayName("False condition keeps else-branch, removes if-branch")
        void falseKeepsElseBranch() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show}",
                    "If content",
                    "#{else}",
                    "Else content",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Before", "Else content", "After");
        }

        @Test
        @DisplayName("Multiple paragraphs in both branches — true")
        void multiParagraphBothBranchesTrue() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Header",
                    "#{if items.size() > 1}",
                    "Table intro paragraph",
                    "Table row 1",
                    "Table row 2",
                    "#{else}",
                    "No table message",
                    "Alternative content",
                    "#{fi}",
                    "Footer"
            );

            Map<String, Object> context = Map.of("items", List.of("a", "b", "c"));
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly(
                    "Header", "Table intro paragraph", "Table row 1", "Table row 2", "Footer"
            );
        }

        @Test
        @DisplayName("Multiple paragraphs in both branches — false")
        void multiParagraphBothBranchesFalse() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Header",
                    "#{if items.size() > 1}",
                    "Table intro paragraph",
                    "Table row 1",
                    "Table row 2",
                    "#{else}",
                    "No table message",
                    "Alternative content",
                    "#{fi}",
                    "Footer"
            );

            Map<String, Object> context = Map.of("items", List.of("a"));
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly(
                    "Header", "No table message", "Alternative content", "Footer"
            );
        }

        @Test
        @DisplayName("If/else with empty if-branch — true keeps nothing between markers")
        void emptyIfBranchTrue() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show}",
                    "#{else}",
                    "Fallback",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Before", "After");
        }

        @Test
        @DisplayName("If/else with empty else-branch — false keeps nothing")
        void emptyElseBranchFalse() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show}",
                    "Content",
                    "#{else}",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Before", "After");
        }

        @Test
        @DisplayName("Two if/else blocks in sequence")
        void twoIfElseBlocksSequential() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Header",
                    "#{if a}",
                    "A-true",
                    "#{else}",
                    "A-false",
                    "#{fi}",
                    "Middle",
                    "#{if b}",
                    "B-true",
                    "#{else}",
                    "B-false",
                    "#{fi}",
                    "Footer"
            );

            Map<String, Object> context = Map.of("a", true, "b", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(2);
            assertThat(extractTexts(content)).containsExactly(
                    "Header", "A-true", "Middle", "B-false", "Footer"
            );
        }

        @Test
        @DisplayName("If/else inside table cell")
        void ifElseInsideTableCell() throws Exception {
            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainPart = docx.getMainDocumentPart();
            mainPart.getContent().clear();

            Tbl table = new Tbl();
            Tr row = new Tr();
            Tc cell = new Tc();

            cell.getContent().add(createParagraph("Cell start"));
            cell.getContent().add(createParagraph("#{if premium}"));
            cell.getContent().add(createParagraph("Premium feature"));
            cell.getContent().add(createParagraph("#{else}"));
            cell.getContent().add(createParagraph("Standard feature"));
            cell.getContent().add(createParagraph("#{fi}"));
            cell.getContent().add(createParagraph("Cell end"));

            row.getContent().add(cell);
            table.getContent().add(row);
            mainPart.getContent().add(table);

            Map<String, Object> context = Map.of("premium", false);

            int processed = DocxConditionalProcessor.processConditionals(
                    mainPart.getContent(), context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(cell.getContent())).containsExactly(
                    "Cell start", "Standard feature", "Cell end"
            );
        }

        @Test
        @DisplayName("Orphan #{else} without #{if} is left in place")
        void orphanElseLeftInPlace() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{else}",
                    "After"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(content, Map.of(), null);

            assertThat(processed).isEqualTo(0);
            assertThat(extractTexts(content)).containsExactly("Before", "#{else}", "After");
        }

        @Test
        @DisplayName("Whitespace around #{else} marker is tolerated")
        void whitespaceAroundElseMarker() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if flag}",
                    "True text",
                    "  #{ else }  ",
                    "False text",
                    "#{fi}"
            );

            Map<String, Object> context = Map.of("flag", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("False text");
        }
    }

    // ==================== MARKER PATTERN TESTS ====================

    @Nested
    @DisplayName("Marker pattern matching")
    class MarkerPatternTests {

        @Test
        @DisplayName("Whitespace around markers is tolerated")
        void whitespaceAroundMarkers() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "  #{  if  show  }  ",
                    "Content",
                    "  #{  fi  }  ",
                    "After"
            );

            Map<String, Object> context = Map.of("show", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Before", "Content", "After");
        }

        @Test
        @DisplayName("IF_PATTERN matches valid markers")
        void ifPatternMatchesValidMarkers() {
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("#{if show}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("#{if debt > 1000}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("#{if status == 'VIP'}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("  #{if show}  ").matches()).isTrue();
        }

        @Test
        @DisplayName("IF_PATTERN rejects invalid markers")
        void ifPatternRejectsInvalid() {
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("#{if}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("#{if }").matches()).isFalse();
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("${if show}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.IF_PATTERN.matcher("Some text #{if show}").matches()).isFalse();
        }

        @Test
        @DisplayName("FI_PATTERN matches valid closing markers")
        void fiPatternMatchesValid() {
            assertThat(DocxConditionalProcessor.FI_PATTERN.matcher("#{fi}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.FI_PATTERN.matcher("  #{fi}  ").matches()).isTrue();
            assertThat(DocxConditionalProcessor.FI_PATTERN.matcher("#{ fi }").matches()).isTrue();
        }

        @Test
        @DisplayName("FI_PATTERN rejects invalid closing markers")
        void fiPatternRejectsInvalid() {
            assertThat(DocxConditionalProcessor.FI_PATTERN.matcher("${fi}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.FI_PATTERN.matcher("Some text #{fi}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.FI_PATTERN.matcher("#{fi} extra").matches()).isFalse();
        }

        @Test
        @DisplayName("ELSE_PATTERN matches valid else markers")
        void elsePatternMatchesValid() {
            assertThat(DocxConditionalProcessor.ELSE_PATTERN.matcher("#{else}").matches()).isTrue();
            assertThat(DocxConditionalProcessor.ELSE_PATTERN.matcher("  #{else}  ").matches()).isTrue();
            assertThat(DocxConditionalProcessor.ELSE_PATTERN.matcher("#{ else }").matches()).isTrue();
        }

        @Test
        @DisplayName("ELSE_PATTERN rejects invalid else markers")
        void elsePatternRejectsInvalid() {
            assertThat(DocxConditionalProcessor.ELSE_PATTERN.matcher("${else}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.ELSE_PATTERN.matcher("Some text #{else}").matches()).isFalse();
            assertThat(DocxConditionalProcessor.ELSE_PATTERN.matcher("#{else} extra").matches()).isFalse();
            assertThat(DocxConditionalProcessor.ELSE_PATTERN.matcher("#{elseif x}").matches()).isFalse();
        }
    }

    // ==================== UNMATCHED MARKER TESTS ====================

    @Nested
    @DisplayName("Unmatched markers")
    class UnmatchedMarkerTests {

        @Test
        @DisplayName("Unmatched #{if} without #{fi} is left in place")
        void unmatchedIfLeftInPlace() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show}",
                    "Content without closing",
                    "After"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(content, Map.of("show", true), null);

            assertThat(processed).isEqualTo(0);
            assertThat(extractTexts(content)).containsExactly(
                    "Before", "#{if show}", "Content without closing", "After"
            );
        }

        @Test
        @DisplayName("Orphan #{fi} is left in place")
        void orphanFiLeftInPlace() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{fi}",
                    "After"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int processed = DocxConditionalProcessor.processConditionals(content, Map.of(), null);

            assertThat(processed).isEqualTo(0);
            assertThat(extractTexts(content)).containsExactly("Before", "#{fi}", "After");
        }
    }

    // ==================== RUN FRAGMENTATION TESTS ====================

    @Nested
    @DisplayName("Run fragmentation handling")
    class RunFragmentationTests {

        @Test
        @DisplayName("Split runs are merged: '#' + '{if show}' → '#{if show}'")
        void splitRunsMerged() throws Exception {
            // Create a paragraph with #{if show} split across two runs
            P paragraph = new P();
            R run1 = new R();
            Text t1 = new Text();
            t1.setValue("#");
            t1.setSpace("preserve");
            run1.getContent().add(t1);

            R run2 = new R();
            Text t2 = new Text();
            t2.setValue("{if show}");
            t2.setSpace("preserve");
            run2.getContent().add(t2);

            paragraph.getContent().add(run1);
            paragraph.getContent().add(run2);

            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            docx.getMainDocumentPart().getContent().clear();
            docx.getMainDocumentPart().getContent().add(createParagraph("Before"));
            docx.getMainDocumentPart().getContent().add(paragraph);
            docx.getMainDocumentPart().getContent().add(createParagraph("Conditional content"));
            docx.getMainDocumentPart().getContent().add(createParagraph("#{fi}"));
            docx.getMainDocumentPart().getContent().add(createParagraph("After"));

            Map<String, Object> context = Map.of("show", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly(
                    "Before", "Conditional content", "After"
            );
        }

        @Test
        @DisplayName("Split runs across three fragments: '#' + '{' + 'if show}'")
        void threeFragmentsMerged() throws Exception {
            P paragraph = new P();

            R run1 = new R();
            Text t1 = new Text();
            t1.setValue("#");
            t1.setSpace("preserve");
            run1.getContent().add(t1);

            R run2 = new R();
            Text t2 = new Text();
            t2.setValue("{");
            t2.setSpace("preserve");
            run2.getContent().add(t2);

            R run3 = new R();
            Text t3 = new Text();
            t3.setValue("if show}");
            t3.setSpace("preserve");
            run3.getContent().add(t3);

            paragraph.getContent().add(run1);
            paragraph.getContent().add(run2);
            paragraph.getContent().add(run3);

            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            docx.getMainDocumentPart().getContent().clear();
            docx.getMainDocumentPart().getContent().add(paragraph);
            docx.getMainDocumentPart().getContent().add(createParagraph("Content"));
            docx.getMainDocumentPart().getContent().add(createParagraph("#{fi}"));

            Map<String, Object> context = Map.of("show", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Content");
        }

        @Test
        @DisplayName("Runs with different formatting are still recognized (paragraph is removed entirely)")
        void differentFormattingStillRecognized() throws Exception {
            P paragraph = new P();

            R run1 = new R();
            RPr rpr1 = new RPr();
            BooleanDefaultTrue bold = new BooleanDefaultTrue();
            bold.setVal(true);
            rpr1.setB(bold);
            run1.setRPr(rpr1);
            Text t1 = new Text();
            t1.setValue("#");
            t1.setSpace("preserve");
            run1.getContent().add(t1);

            R run2 = new R();
            Text t2 = new Text();
            t2.setValue("{if show}");
            t2.setSpace("preserve");
            run2.getContent().add(t2);

            paragraph.getContent().add(run1);
            paragraph.getContent().add(run2);

            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            docx.getMainDocumentPart().getContent().clear();
            docx.getMainDocumentPart().getContent().add(createParagraph("Before"));
            docx.getMainDocumentPart().getContent().add(paragraph);
            docx.getMainDocumentPart().getContent().add(createParagraph("Content"));
            docx.getMainDocumentPart().getContent().add(createParagraph("#{fi}"));
            docx.getMainDocumentPart().getContent().add(createParagraph("After"));

            List<Object> content = docx.getMainDocumentPart().getContent();

            int processed = DocxConditionalProcessor.processConditionals(content, Map.of("show", true), null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(content)).containsExactly("Before", "Content", "After");
        }
    }

    // ==================== TABLE CELL CONDITIONAL TESTS ====================

    @Nested
    @DisplayName("Conditionals inside table cells")
    class TableCellConditionalTests {

        @Test
        @DisplayName("Conditional inside a table cell")
        void conditionalInsideTableCell() throws Exception {
            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainPart = docx.getMainDocumentPart();
            mainPart.getContent().clear();

            Tbl table = new Tbl();
            Tr row = new Tr();
            Tc cell = new Tc();

            cell.getContent().add(createParagraph("Cell before"));
            cell.getContent().add(createParagraph("#{if show_detail}"));
            cell.getContent().add(createParagraph("Detail info"));
            cell.getContent().add(createParagraph("#{fi}"));
            cell.getContent().add(createParagraph("Cell after"));

            row.getContent().add(cell);
            table.getContent().add(row);
            mainPart.getContent().add(table);

            Map<String, Object> context = Map.of("show_detail", false);

            int processed = DocxConditionalProcessor.processConditionals(
                    mainPart.getContent(), context, null);

            assertThat(processed).isEqualTo(1);
            assertThat(extractTexts(cell.getContent())).containsExactly("Cell before", "Cell after");
        }
    }

    // ==================== MULTI-PARAGRAPH CONTENT TESTS ====================

    @Nested
    @DisplayName("Multi-paragraph conditional content")
    class MultiParagraphContentTests {

        @Test
        @DisplayName("Multiple paragraphs between markers removed when false")
        void multiParagraphRemovedWhenFalse() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show_section}",
                    "Paragraph 1",
                    "Paragraph 2",
                    "Paragraph 3",
                    "Paragraph 4",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show_section", false);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Before", "After");
        }

        @Test
        @DisplayName("Multiple paragraphs between markers kept when true")
        void multiParagraphKeptWhenTrue() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Before",
                    "#{if show_section}",
                    "Paragraph 1",
                    "Paragraph 2",
                    "Paragraph 3",
                    "#{fi}",
                    "After"
            );

            Map<String, Object> context = Map.of("show_section", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly(
                    "Before", "Paragraph 1", "Paragraph 2", "Paragraph 3", "After"
            );
        }

        @Test
        @DisplayName("Adjacent conditional blocks with no content between")
        void adjacentBlocks() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if show_a}",
                    "Block A",
                    "#{fi}",
                    "#{if show_b}",
                    "Block B",
                    "#{fi}"
            );

            Map<String, Object> context = Map.of("show_a", false, "show_b", true);
            List<Object> content = docx.getMainDocumentPart().getContent();

            DocxConditionalProcessor.processConditionals(content, context, null);

            assertThat(extractTexts(content)).containsExactly("Block B");
        }
    }

    // ==================== EVALUATE CONDITION TESTS ====================

    @Nested
    @DisplayName("evaluateCondition")
    class EvaluateConditionTests {

        @Test
        @DisplayName("Boolean true expression")
        void booleanTrueExpression() {
            Map<String, Object> context = Map.of("x", 10);
            assertThat(DocxConditionalProcessor.evaluateCondition("x > 5", context, null)).isTrue();
        }

        @Test
        @DisplayName("Boolean false expression")
        void booleanFalseExpression() {
            Map<String, Object> context = Map.of("x", 3);
            assertThat(DocxConditionalProcessor.evaluateCondition("x > 5", context, null)).isFalse();
        }

        @Test
        @DisplayName("String comparison returns boolean")
        void stringComparisonReturnsBoolean() {
            Map<String, Object> context = Map.of("role", "admin");
            assertThat(DocxConditionalProcessor.evaluateCondition("role == 'admin'", context, null)).isTrue();
        }

        @Test
        @DisplayName("Failed expression evaluation treated as false")
        void failedExpressionTreatedAsFalse() {
            assertThat(DocxConditionalProcessor.evaluateCondition(
                    "nonexistent_variable > 5", Map.of(), null)).isFalse();
        }
    }

    // ==================== FIND CLOSING MARKERS TESTS ====================

    @Nested
    @DisplayName("findClosingMarkers / findClosingFi")
    class FindClosingMarkersTests {

        @Test
        @DisplayName("Finds #{fi} at correct index")
        void findsClosingFi() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if show}",
                    "Content",
                    "#{fi}",
                    "After"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int fiIndex = DocxConditionalProcessor.findClosingFi(content, 1);

            assertThat(fiIndex).isEqualTo(2);
        }

        @Test
        @DisplayName("Returns -1 when no #{fi} found")
        void returnsNegativeOneWhenNotFound() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if show}",
                    "Content",
                    "No closing marker"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int fiIndex = DocxConditionalProcessor.findClosingFi(content, 1);

            assertThat(fiIndex).isEqualTo(-1);
        }

        @Test
        @DisplayName("findClosingMarkers returns elseIndex and fiIndex")
        void findsBothElseAndFi() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if show}",
                    "True content",
                    "#{else}",
                    "False content",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int[] markers = DocxConditionalProcessor.findClosingMarkers(content, 1);

            assertThat(markers[0]).isEqualTo(2); // #{else}
            assertThat(markers[1]).isEqualTo(4); // #{fi}
        }

        @Test
        @DisplayName("findClosingMarkers returns -1 for else when no #{else}")
        void noElseReturnsMinusOne() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if show}",
                    "Content",
                    "#{fi}"
            );

            List<Object> content = docx.getMainDocumentPart().getContent();
            int[] markers = DocxConditionalProcessor.findClosingMarkers(content, 1);

            assertThat(markers[0]).isEqualTo(-1); // no #{else}
            assertThat(markers[1]).isEqualTo(2);  // #{fi}
        }
    }

    // ==================== extractConditionalVariables() ====================

    @Nested
    @DisplayName("extractConditionalVariables() — variable extraction from #{if} markers")
    class ExtractConditionalVariablesTests {

        @Test
        @DisplayName("simple variable in #{if}")
        void simpleVariable() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if active}",
                    "Content",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("active");
        }

        @Test
        @DisplayName("comparison expression extracts variable")
        void comparisonExpression() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if debt > 1000}",
                    "Overdue notice",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("debt");
        }

        @Test
        @DisplayName("logical AND extracts both variables")
        void logicalAndExpression() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if active && verified}",
                    "Verified content",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("active", "verified");
        }

        @Test
        @DisplayName("negation extracts variable")
        void negationExpression() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if !expired}",
                    "Valid content",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("expired");
        }

        @Test
        @DisplayName("ternary expression inside #{if} extracts all variables")
        void ternaryExpression() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if score > threshold}",
                    "Above threshold",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("score", "threshold");
        }

        @Test
        @DisplayName("multiple #{if} blocks collect all variables")
        void multipleIfBlocks() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if active}",
                    "Block 1",
                    "#{fi}",
                    "#{if debt > 1000}",
                    "Block 2",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("active", "debt");
        }

        @Test
        @DisplayName("no conditional blocks returns empty set")
        void noConditionals() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "Normal paragraph",
                    "Another paragraph"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).isEmpty();
        }

        @Test
        @DisplayName("string comparison — keyword not extracted, variable is")
        void stringComparison() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if status == 'active'}",
                    "Active content",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("status");
        }

        @Test
        @DisplayName("complex expression with multiple operators")
        void complexExpression() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if age >= 18 && income > 50000 && !blacklisted}",
                    "Eligible",
                    "#{fi}"
            );

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    docx.getMainDocumentPart().getContent());

            assertThat(vars).containsExactly("age", "income", "blacklisted");
        }

        @Test
        @DisplayName("conditional inside table cell")
        void conditionalInTableCell() throws Exception {
            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainPart = docx.getMainDocumentPart();
            mainPart.getContent().clear();

            // Build: Tbl > Tr > Tc > [#{if debt > 1000}, Content, #{fi}]
            Tbl table = new Tbl();
            Tr row = new Tr();
            Tc cell = new Tc();
            cell.getContent().add(createParagraph("#{if debt > 1000}"));
            cell.getContent().add(createParagraph("Overdue"));
            cell.getContent().add(createParagraph("#{fi}"));
            row.getContent().add(cell);
            table.getContent().add(row);
            mainPart.getContent().add(table);

            Set<String> vars = DocxConditionalProcessor.extractConditionalVariables(
                    mainPart.getContent());

            assertThat(vars).containsExactly("debt");
        }
    }
}
