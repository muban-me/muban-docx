package me.muban.docx;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.expression.MapAccessor;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.Expression;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.SimpleEvaluationContext;

import java.util.Collections;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * Evaluates SpEL (Spring Expression Language) expressions against a map-based context.
 *
 * <p>Uses a security-sandboxed {@link SimpleEvaluationContext} with {@link MapAccessor}
 * that allows property access on maps and method calls on standard types, but prevents
 * arbitrary class instantiation or static method invocation.
 *
 * <p>Results are formatted using {@link LocaleUtils#formatValue(Object, Locale)} when
 * a locale is provided (numbers get locale-specific grouping/decimal separators).
 *
 * <p>On evaluation failure, returns the original placeholder text {@code ${expression}}
 * so unresolvable expressions remain visible in the output.
 */
public class DocxExpressionEvaluator {

    private static final Logger log = LoggerFactory.getLogger(DocxExpressionEvaluator.class);
    private static final ExpressionParser PARSER = new SpelExpressionParser();
    private static final Pattern SIMPLE_KEY_PATTERN = Pattern.compile("[a-zA-Z_][a-zA-Z0-9_.]*");

    private DocxExpressionEvaluator() {}

    /**
     * Check whether a placeholder body is an expression (contains operators, method calls, etc.)
     * or a simple dotted key like {@code customer_name} or {@code address.city}.
     *
     * @param placeholderBody the text between {@code ${} and {@code }}
     * @return true if the body contains expression syntax beyond a simple key
     */
    public static boolean isExpression(String placeholderBody) {
        if (placeholderBody == null || placeholderBody.isBlank()) {
            return false;
        }
        return !SIMPLE_KEY_PATTERN.matcher(placeholderBody.trim()).matches();
    }

    /**
     * Evaluate a SpEL expression against the given context.
     *
     * @param expression the expression to evaluate (without {@code ${} / {@code }} delimiters)
     * @param rawContext  the context map — variables are accessible as properties
     * @param locale     optional locale for number formatting, or null
     * @return the formatted result, or {@code "${expression}"} on failure
     */
    public static String evaluate(String expression, Map<String, Object> rawContext, Locale locale) {
        if (rawContext == null) {
            rawContext = Collections.emptyMap();
        }
        try {
            EvaluationContext evalContext = SimpleEvaluationContext
                    .forPropertyAccessors(new MapAccessor())
                    .withInstanceMethods()
                    .build();
            Expression spelExpression = PARSER.parseExpression(expression);
            Object result = spelExpression.getValue(evalContext, rawContext);
            if (result == null) {
                return "";
            }
            return LocaleUtils.formatValue(result, locale);
        } catch (Exception e) {
            log.warn("SpEL expression evaluation failed for '{}': {}", expression, e.getMessage());
            return "${" + expression + "}";
        }
    }
}
