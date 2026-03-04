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
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
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

    /**
     * Pattern matching identifiers in SpEL expressions: sequences of letters, digits,
     * underscores and dots that start with a letter or underscore.
     */
    private static final Pattern IDENTIFIER_PATTERN =
            Pattern.compile("[a-zA-Z_][a-zA-Z0-9_.]*");

    /**
     * SpEL keywords and literals that should not be treated as variable references.
     */
    private static final Set<String> SPEL_KEYWORDS = Set.of(
            "true", "false", "null",
            "new", "instanceof", "T",
            "and", "or", "not",
            "eq", "ne", "lt", "gt", "le", "ge",
            "div", "mod",
            "matches", "between"
    );

    /**
     * Pattern to strip string literals (single- and double-quoted) from expressions
     * before scanning for identifiers. Handles escaped quotes.
     */
    private static final Pattern STRING_LITERAL_PATTERN =
            Pattern.compile("'(?:[^'\\\\]|\\\\.)*'|\"(?:[^\"\\\\]|\\\\.)*\"");

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

    /**
     * Extract variable references from a SpEL expression string.
     *
     * <p>Uses a regex heuristic: finds all identifier tokens, strips those inside
     * string literals, removes SpEL keywords ({@code true}, {@code false}, {@code null},
     * {@code and}, {@code or}, etc.), and removes method names (identifiers immediately
     * followed by {@code (}).
     *
     * <p>Examples:
     * <ul>
     *   <li>{@code "available > 10 ? \"many\" : small_amount_desc"}
     *       → {@code ["available", "small_amount_desc"]}</li>
     *   <li>{@code "debt > 1000"} → {@code ["debt"]}</li>
     *   <li>{@code "active && verified"} → {@code ["active", "verified"]}</li>
     *   <li>{@code "name.toUpperCase()"} → {@code []} (method call — variable discovery
     *       relies on other occurrences of {@code name} in the template)</li>
     *   <li>{@code "price * quantity"} → {@code ["price", "quantity"]}</li>
     *   <li>{@code "gender == 'F' ? 'a.png' : 'b.png'"} → {@code ["gender"]}</li>
     * </ul>
     *
     * <p><b>Method calls are skipped entirely</b> — {@code items.size()} does not extract
     * {@code items} because data arrays are already discovered through their field references
     * (e.g., {@code items.name}, {@code items.price}) in table placeholders.
     *
     * @param expression the SpEL expression (without {@code ${}} or {@code #{if}} delimiters)
     * @return ordered set of variable names found, never null
     */
    public static Set<String> extractVariableReferences(String expression) {
        if (expression == null || expression.isBlank()) {
            return Collections.emptySet();
        }

        // Step 1: Remove string literals so we don't pick up identifiers inside quotes
        String stripped = STRING_LITERAL_PATTERN.matcher(expression).replaceAll(" ");

        // Step 2: Find all identifier tokens
        Set<String> variables = new LinkedHashSet<>();
        Matcher matcher = IDENTIFIER_PATTERN.matcher(stripped);

        while (matcher.find()) {
            String token = matcher.group();
            int end = matcher.end();

            // Skip method calls entirely: identifier immediately followed by '('
            // e.g., "items.size()" or "name.toUpperCase()" — the object part is NOT
            // extracted because we cannot distinguish scalar variables (name) from
            // data arrays (items) at this level. Arrays are discovered via their
            // field placeholders (items.name, items.price) in table rows.
            if (end < stripped.length() && stripped.charAt(end) == '(') {
                continue;
            }

            // Skip SpEL keywords
            if (SPEL_KEYWORDS.contains(token)) {
                continue;
            }

            // Skip pure numeric-looking tokens after a dot (e.g., shouldn't happen with our
            // pattern, but guard against "10.0" being partially matched)
            variables.add(token);
        }

        return variables;
    }
}
