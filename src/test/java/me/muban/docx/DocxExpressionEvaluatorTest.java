package me.muban.docx;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.*;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link DocxExpressionEvaluator} — SpEL expression evaluation
 * in DOCX template placeholders.
 */
@DisplayName("DocxExpressionEvaluator")
class DocxExpressionEvaluatorTest {

    // ==================== isExpression() ====================

    @Nested
    @DisplayName("isExpression() — detection heuristic")
    class IsExpressionTests {

        @Test
        @DisplayName("simple key names are NOT expressions")
        void simpleKeysAreNotExpressions() {
            assertThat(DocxExpressionEvaluator.isExpression("recipientName")).isFalse();
            assertThat(DocxExpressionEvaluator.isExpression("caseNumber")).isFalse();
            assertThat(DocxExpressionEvaluator.isExpression("_private")).isFalse();
            assertThat(DocxExpressionEvaluator.isExpression("my_var")).isFalse();
        }

        @Test
        @DisplayName("dot-notation keys are NOT expressions")
        void dotNotationKeysAreNotExpressions() {
            assertThat(DocxExpressionEvaluator.isExpression("address.city")).isFalse();
            assertThat(DocxExpressionEvaluator.isExpression("items.name")).isFalse();
            assertThat(DocxExpressionEvaluator.isExpression("company.address.zipCode")).isFalse();
        }

        @Test
        @DisplayName("ternary expressions ARE expressions")
        void ternaryExpressionsAreExpressions() {
            assertThat(DocxExpressionEvaluator.isExpression(
                    "items.size() > 3 ? \"a lot\" : \"a few\"")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression(
                    "\"female\".equals(gender) ? \"Mrs.\" : \"Mr.\"")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression(
                    "age >= 18 ? \"adult\" : \"minor\"")).isTrue();
        }

        @Test
        @DisplayName("method calls ARE expressions")
        void methodCallsAreExpressions() {
            assertThat(DocxExpressionEvaluator.isExpression("items.size()")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression("name.toUpperCase()")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression("name.length()")).isTrue();
        }

        @Test
        @DisplayName("arithmetic expressions ARE expressions")
        void arithmeticExpressionsAreExpressions() {
            assertThat(DocxExpressionEvaluator.isExpression("count + 1")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression("price * qty")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression("total - discount")).isTrue();
        }

        @Test
        @DisplayName("comparison expressions ARE expressions")
        void comparisonExpressionsAreExpressions() {
            assertThat(DocxExpressionEvaluator.isExpression("count > 0")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression("status == 'active'")).isTrue();
            assertThat(DocxExpressionEvaluator.isExpression("count != 0")).isTrue();
        }

        @Test
        @DisplayName("null and blank input are NOT expressions")
        void nullAndBlankAreNotExpressions() {
            assertThat(DocxExpressionEvaluator.isExpression(null)).isFalse();
            assertThat(DocxExpressionEvaluator.isExpression("")).isFalse();
            assertThat(DocxExpressionEvaluator.isExpression("   ")).isFalse();
        }
    }

    // ==================== evaluate() — ternary expressions ====================

    @Nested
    @DisplayName("evaluate() — ternary conditional expressions")
    class TernaryExpressionTests {

        @Test
        @DisplayName("ternary with numeric comparison — true branch")
        void ternaryNumericTrue() {
            Map<String, Object> ctx = Map.of("items", List.of("a", "b", "c", "d", "e"));

            String result = DocxExpressionEvaluator.evaluate(
                    "items.size() > 3 ? \"a lot\" : \"a few\"", ctx, null);

            assertThat(result).isEqualTo("a lot");
        }

        @Test
        @DisplayName("ternary with numeric comparison — false branch")
        void ternaryNumericFalse() {
            Map<String, Object> ctx = Map.of("items", List.of("a", "b"));

            String result = DocxExpressionEvaluator.evaluate(
                    "items.size() > 3 ? \"a lot\" : \"a few\"", ctx, null);

            assertThat(result).isEqualTo("a few");
        }

        @Test
        @DisplayName("ternary with string equality — true branch")
        void ternaryStringEqualityTrue() {
            Map<String, Object> ctx = Map.of("gender", "female");

            String result = DocxExpressionEvaluator.evaluate(
                    "\"female\".equals(gender) ? \"Mrs.\" : \"Mr.\"", ctx, null);

            assertThat(result).isEqualTo("Mrs.");
        }

        @Test
        @DisplayName("ternary with string equality — false branch")
        void ternaryStringEqualityFalse() {
            Map<String, Object> ctx = Map.of("gender", "male");

            String result = DocxExpressionEvaluator.evaluate(
                    "\"female\".equals(gender) ? \"Mrs.\" : \"Mr.\"", ctx, null);

            assertThat(result).isEqualTo("Mr.");
        }

        @Test
        @DisplayName("ternary with integer comparison")
        void ternaryIntegerComparison() {
            Map<String, Object> ctx = Map.of("age", 21);

            String result = DocxExpressionEvaluator.evaluate(
                    "age >= 18 ? \"adult\" : \"minor\"", ctx, null);

            assertThat(result).isEqualTo("adult");
        }

        @Test
        @DisplayName("ternary with integer comparison — minor path")
        void ternaryIntegerComparisonMinor() {
            Map<String, Object> ctx = Map.of("age", 15);

            String result = DocxExpressionEvaluator.evaluate(
                    "age >= 18 ? \"adult\" : \"minor\"", ctx, null);

            assertThat(result).isEqualTo("minor");
        }

        @Test
        @DisplayName("ternary for singular/plural")
        void ternarySingularPlural() {
            Map<String, Object> ctx1 = Map.of("count", 1);
            Map<String, Object> ctx5 = Map.of("count", 5);

            assertThat(DocxExpressionEvaluator.evaluate(
                    "count == 1 ? \"item\" : \"items\"", ctx1, null))
                    .isEqualTo("item");

            assertThat(DocxExpressionEvaluator.evaluate(
                    "count == 1 ? \"item\" : \"items\"", ctx5, null))
                    .isEqualTo("items");
        }

        @Test
        @DisplayName("ternary with null check using Elvis operator")
        void ternaryElvisOperator() {
            Map<String, Object> ctx = new HashMap<>();
            ctx.put("nickname", null);

            String result = DocxExpressionEvaluator.evaluate(
                    "nickname ?: \"N/A\"", ctx, null);

            assertThat(result).isEqualTo("N/A");
        }

        @Test
        @DisplayName("ternary with empty list check")
        void ternaryEmptyListCheck() {
            Map<String, Object> ctx = Map.of("items", List.of());

            String result = DocxExpressionEvaluator.evaluate(
                    "items.isEmpty() ? \"No items\" : \"Has items\"", ctx, null);

            assertThat(result).isEqualTo("No items");
        }
    }

    // ==================== evaluate() — method calls ====================

    @Nested
    @DisplayName("evaluate() — method calls and property access")
    class MethodCallTests {

        @Test
        @DisplayName("list size method call")
        void listSizeMethod() {
            Map<String, Object> ctx = Map.of("items", List.of("a", "b", "c"));

            String result = DocxExpressionEvaluator.evaluate("items.size()", ctx, null);

            assertThat(result).isEqualTo("3");
        }

        @Test
        @DisplayName("string toUpperCase method call")
        void stringToUpperCase() {
            Map<String, Object> ctx = Map.of("name", "john");

            String result = DocxExpressionEvaluator.evaluate("name.toUpperCase()", ctx, null);

            assertThat(result).isEqualTo("JOHN");
        }

        @Test
        @DisplayName("string length method call")
        void stringLength() {
            Map<String, Object> ctx = Map.of("name", "Hello");

            String result = DocxExpressionEvaluator.evaluate("name.length()", ctx, null);

            assertThat(result).isEqualTo("5");
        }

        @Test
        @DisplayName("nested map access via dot notation")
        void nestedMapAccess() {
            Map<String, Object> ctx = Map.of(
                    "address", Map.of("city", "Warsaw", "country", "Poland"));

            String result = DocxExpressionEvaluator.evaluate("address.city", ctx, null);

            assertThat(result).isEqualTo("Warsaw");
        }

        @Test
        @DisplayName("nested map with ternary")
        void nestedMapWithTernary() {
            Map<String, Object> ctx = Map.of(
                    "address", Map.of("city", "Warsaw"));

            String result = DocxExpressionEvaluator.evaluate(
                    "\"Warsaw\".equals(address.city) ? \"local\" : \"remote\"", ctx, null);

            assertThat(result).isEqualTo("local");
        }

        @Test
        @DisplayName("arithmetic expression")
        void arithmeticExpression() {
            Map<String, Object> ctx = Map.of("price", 10, "qty", 3);

            String result = DocxExpressionEvaluator.evaluate("price * qty", ctx, null);

            assertThat(result).isEqualTo("30");
        }
    }

    // ==================== evaluate() — locale formatting ====================

    @Nested
    @DisplayName("evaluate() — locale-aware result formatting")
    class LocaleFormattingTests {

        @Test
        @DisplayName("numeric result formatted with Polish locale")
        void numericResultWithPolishLocale() {
            Map<String, Object> ctx = Map.of("price", 1000, "qty", 3);

            String result = DocxExpressionEvaluator.evaluate(
                    "price * qty", ctx, new Locale("pl", "PL"));

            // Polish locale uses non-breaking space as grouping separator
            assertThat(result).contains("3");
            assertThat(result).contains("000");
        }

        @Test
        @DisplayName("string result not affected by locale")
        void stringResultIgnoresLocale() {
            Map<String, Object> ctx = Map.of("name", "John");

            String result = DocxExpressionEvaluator.evaluate(
                    "name.toUpperCase()", ctx, new Locale("pl", "PL"));

            assertThat(result).isEqualTo("JOHN");
        }
    }

    // ==================== evaluate() — error handling ====================

    @Nested
    @DisplayName("evaluate() — error handling and edge cases")
    class ErrorHandlingTests {

        @Test
        @DisplayName("invalid expression returns original placeholder")
        void invalidExpressionReturnsOriginal() {
            Map<String, Object> ctx = Map.of("name", "John");

            String result = DocxExpressionEvaluator.evaluate(
                    "invalid %%% expression", ctx, null);

            assertThat(result).isEqualTo("${invalid %%% expression}");
        }

        @Test
        @DisplayName("missing variable returns original placeholder")
        void missingVariableReturnsOriginal() {
            Map<String, Object> ctx = Map.of("name", "John");

            String result = DocxExpressionEvaluator.evaluate(
                    "nonexistent.size() > 0 ? \"yes\" : \"no\"", ctx, null);

            // Should return original since 'nonexistent' is not in context
            assertThat(result).startsWith("${");
        }

        @Test
        @DisplayName("null result returns empty string")
        void nullResultReturnsEmpty() {
            Map<String, Object> ctx = new HashMap<>();
            ctx.put("value", null);

            String result = DocxExpressionEvaluator.evaluate("value", ctx, null);

            assertThat(result).isEmpty();
        }

        @Test
        @DisplayName("null raw context treated as empty map")
        void nullRawContext() {
            String result = DocxExpressionEvaluator.evaluate(
                    "1 + 1", null, null);

            assertThat(result).isEqualTo("2");
        }

        @Test
        @DisplayName("empty raw context — literal expressions still work")
        void emptyContext() {
            String result = DocxExpressionEvaluator.evaluate(
                    "1 > 0 ? \"yes\" : \"no\"", Map.of(), null);

            assertThat(result).isEqualTo("yes");
        }
    }

    // ==================== Security ====================

    @Nested
    @DisplayName("Security — blocked operations")
    class SecurityTests {

        @Test
        @DisplayName("type references (T()) are blocked — returns original placeholder")
        void typeReferencesBlocked() {
            Map<String, Object> ctx = Map.of();

            String result = DocxExpressionEvaluator.evaluate(
                    "T(java.lang.Runtime).getRuntime().exec('calc')", ctx, null);

            // Should fail and return original placeholder
            assertThat(result).startsWith("${");
        }

        @Test
        @DisplayName("constructor calls (new) are blocked — returns original placeholder")
        void constructorCallsBlocked() {
            Map<String, Object> ctx = Map.of();

            String result = DocxExpressionEvaluator.evaluate(
                    "new java.lang.ProcessBuilder('calc').start()", ctx, null);

            assertThat(result).startsWith("${");
        }

        @Test
        @DisplayName("bean references (@) are blocked — returns original placeholder")
        void beanReferencesBlocked() {
            Map<String, Object> ctx = Map.of();

            String result = DocxExpressionEvaluator.evaluate(
                    "@systemProperties", ctx, null);

            assertThat(result).startsWith("${");
        }
    }

    // ==================== buildRawContext integration ====================

    @Nested
    @DisplayName("Integration with DocxContextBuilder.buildRawContext()")
    class RawContextIntegrationTests {

        @Test
        @DisplayName("raw context preserves list type for size() calls")
        void rawContextPreservesListType() {
            Map<String, Object> data = Map.of(
                    "items", List.of(
                            Map.of("name", "A"),
                            Map.of("name", "B"),
                            Map.of("name", "C")));

            Map<String, Object> rawCtx = DocxContextBuilder.buildRawContext(data);

            String result = DocxExpressionEvaluator.evaluate(
                    "items.size() > 2 ? \"many\" : \"few\"", rawCtx, null);

            assertThat(result).isEqualTo("many");
        }

        @Test
        @DisplayName("raw context preserves numeric type for comparison")
        void rawContextPreservesNumericType() {
            // In the library, numeric values come via the data map
            Map<String, Object> data = Map.of("age", 25);

            Map<String, Object> rawCtx = DocxContextBuilder.buildRawContext(data);

            String result = DocxExpressionEvaluator.evaluate(
                    "age >= 18 ? \"adult\" : \"minor\"", rawCtx, null);

            assertThat(result).isEqualTo("adult");
        }

        @Test
        @DisplayName("raw context merges parameters and data — data wins on conflict")
        void rawContextMergesWithDataPrecedence() {
            Map<String, String> params = Map.of("status", "draft");
            Map<String, Object> data = Map.of("status", "published");

            Map<String, Object> rawCtx = DocxContextBuilder.buildRawContext(params, data);

            String result = DocxExpressionEvaluator.evaluate(
                    "\"published\".equals(status) ? \"live\" : \"preview\"", rawCtx, null);

            assertThat(result).isEqualTo("live");
        }
    }

    // ==================== extractVariableReferences() ====================

    @Nested
    @DisplayName("extractVariableReferences() — variable extraction from SpEL")
    class ExtractVariableReferencesTests {

        @Test
        @DisplayName("simple variable name")
        void simpleVariable() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("name"))
                    .containsExactly("name");
        }

        @Test
        @DisplayName("dotted variable name")
        void dottedVariable() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("address.city"))
                    .containsExactly("address.city");
        }

        @Test
        @DisplayName("comparison expression extracts both sides")
        void comparisonExpression() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("available > 10"))
                    .containsExactly("available");
        }

        @Test
        @DisplayName("ternary with variable in else branch")
        void ternaryWithVariableInElse() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "available > 10 ? \"many\" : small_amount_desc"))
                    .containsExactly("available", "small_amount_desc");
        }

        @Test
        @DisplayName("ternary with string literals only — no false positives")
        void ternaryWithStringLiteralsOnly() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "gender == 'F' ? 'Mrs.' : 'Mr.'"))
                    .containsExactly("gender");
        }

        @Test
        @DisplayName("logical AND with two variables")
        void logicalAndExpression() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("active && verified"))
                    .containsExactly("active", "verified");
        }

        @Test
        @DisplayName("logical OR with two variables")
        void logicalOrExpression() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("active || backup"))
                    .containsExactly("active", "backup");
        }

        @Test
        @DisplayName("negation extracts variable")
        void negationExpression() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("!expired"))
                    .containsExactly("expired");
        }

        @Test
        @DisplayName("arithmetic extracts both operands")
        void arithmeticExpression() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("price * quantity"))
                    .containsExactly("price", "quantity");
        }

        @Test
        @DisplayName("method call — entire token skipped")
        void methodCallExpression() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("name.toUpperCase()"))
                    .isEmpty();
        }

        @Test
        @DisplayName("chained method call — entire token skipped")
        void chainedMethodCall() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("name.trim().toUpperCase()"))
                    .isEmpty();
        }

        @Test
        @DisplayName("string concatenation with variables")
        void stringConcatenation() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "firstName + ' ' + lastName"))
                    .containsExactly("firstName", "lastName");
        }

        @Test
        @DisplayName("complex ternary with comparison and variables")
        void complexTernary() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "debt > 1000 ? overdue_msg : normal_msg"))
                    .containsExactly("debt", "overdue_msg", "normal_msg");
        }

        @Test
        @DisplayName("SpEL keyword 'true' is not extracted as variable")
        void spelKeywordTrue() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("active == true"))
                    .containsExactly("active");
        }

        @Test
        @DisplayName("SpEL keyword 'false' is not extracted as variable")
        void spelKeywordFalse() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("active == false"))
                    .containsExactly("active");
        }

        @Test
        @DisplayName("SpEL keyword 'null' is not extracted as variable")
        void spelKeywordNull() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("value != null"))
                    .containsExactly("value");
        }

        @Test
        @DisplayName("SpEL textual operators (and, or, not) are not extracted")
        void spelTextualOperators() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "active and verified or not expired"))
                    .containsExactly("active", "verified", "expired");
        }

        @Test
        @DisplayName("null input returns empty set")
        void nullInput() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(null)).isEmpty();
        }

        @Test
        @DisplayName("blank input returns empty set")
        void blankInput() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("   ")).isEmpty();
        }

        @Test
        @DisplayName("pure string literal returns empty set")
        void pureStringLiteral() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("'hello world'")).isEmpty();
        }

        @Test
        @DisplayName("numeric comparison with decimal — no false identifiers")
        void numericComparison() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences("price > 10.50"))
                    .containsExactly("price");
        }

        @Test
        @DisplayName("double-quoted strings inside expression are excluded")
        void doubleQuotedStrings() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "status == \"active\" ? \"yes\" : \"no\""))
                    .containsExactly("status");
        }

        @Test
        @DisplayName("image-style conditional expression")
        void imageConditionalExpression() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "gender == 'F' ? 'assets/female.png' : 'assets/male.png'"))
                    .containsExactly("gender");
        }

        @Test
        @DisplayName("mixed path with embedded expression")
        void mixedPathExpression() {
            // "assets/${dept}/stamp.png" — when extracted from ${...} the body is just "dept"
            // but for the dept variable in a larger expression:
            assertThat(DocxExpressionEvaluator.extractVariableReferences("dept"))
                    .containsExactly("dept");
        }

        @Test
        @DisplayName("equals method call on string literal")
        void equalsMethodOnLiteral() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "\"active\".equals(status)"))
                    .containsExactly("status");
        }

        @Test
        @DisplayName("multiple comparisons with logical operators")
        void multipleComparisons() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "age >= 18 && income > 50000 && !blacklisted"))
                    .containsExactly("age", "income", "blacklisted");
        }

        @Test
        @DisplayName("no duplicates in result")
        void noDuplicates() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "total > 0 ? total : 0"))
                    .containsExactly("total");
        }

        @Test
        @DisplayName("size() method call on collection — entire token skipped")
        void sizeMethodCall() {
            assertThat(DocxExpressionEvaluator.extractVariableReferences(
                    "items.size() > 3 ? 'many' : 'few'"))
                    .isEmpty();
        }
    }
}
