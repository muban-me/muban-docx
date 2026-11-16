package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static me.muban.docx.DocxTestSupport.createDocx;
import static me.muban.docx.DocxTestSupport.extractTexts;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for the expression side of {@link DocxConditionalProcessor}: how a {@code #{if}} condition
 * is evaluated, and how its result is coerced to a boolean.
 */
@DisplayName("DocxConditionalProcessor — expression evaluation")
class DocxConditionalEvaluationTest {

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
}
