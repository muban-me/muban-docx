package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Tr;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.Set;

import static me.muban.docx.DocxTestSupport.createDocx;
import static me.muban.docx.DocxTestSupport.createParagraph;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for {@link DocxConditionalProcessor#extractConditionalVariables(java.util.List)} —
 * discovering the variables a template's conditional expressions depend on, so that template
 * upload can register the right parameters.
 */
@DisplayName("DocxConditionalProcessor — variable extraction")
class DocxConditionalVariablesTest {

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
