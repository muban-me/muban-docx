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

    /**
     * SpEL variables that always exist and therefore need no declaration.
     */
    private static final Set<String> PREDEFINED_VARIABLES = Set.of("root", "this");

    /**
     * Pattern for a variable reference: {@code #name}, with the name captured up to the first dot so
     * {@code #fmt.date(...)} checks {@code fmt} rather than {@code fmt.date}.
     */
    private static final Pattern VARIABLE_REFERENCE_PATTERN =
            Pattern.compile("#([a-zA-Z_][a-zA-Z0-9_]*)");

    private DocxExpressionEvaluator() {}

    /**
     * Sentinel returned by {@link #evaluateRaw} when an expression cannot be evaluated.
     *
     * <p>Distinguishes a genuine {@code null} result (which is a valid value and may be bound to a
     * template variable) from a failure (which must leave the variable undefined).
     */
    static final Object UNRESOLVED = new Object();

    /**
     * Outcome of an expression evaluation: either a value (possibly {@code null}) or a failure.
     */
    private record Result(boolean resolved, Object value) {

        static Result of(Object value) {
            return new Result(true, value);
        }

        static Result failure() {
            return new Result(false, null);
        }
    }

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
        return evaluate(expression, rawContext, locale, Collections.emptyMap());
    }

    /**
     * Evaluate a SpEL expression with additional variables reachable through the {@code #} prefix.
     *
     * <p>Variables are how the engine exposes helpers to templates without weakening the sandbox:
     * the document locale's {@link DocxFormatHelper} is available as {@code #fmt}, so a template can
     * write {@code ${#fmt.date(document_date, 'dd.MM.yyyy')}} even though static method calls are
     * forbidden.
     *
     * @param expression the expression to evaluate (without {@code ${} / {@code }} delimiters)
     * @param rawContext  the context map — variables are accessible as properties
     * @param locale     optional locale for number formatting, or null
     * @param variables  extra variables exposed as {@code #name}; may be {@code null}
     * @return the formatted result, or {@code "${expression}"} on failure
     */
    public static String evaluate(String expression, Map<String, Object> rawContext, Locale locale,
                                  Map<String, Object> variables) {
        Result result = evaluateInternal(expression, rawContext, locale, variables);
        if (!result.resolved()) {
            return "${" + expression + "}";
        }
        if (result.value() == null) {
            return "";
        }
        return LocaleUtils.formatValue(result.value(), locale);
    }

    /**
     * Evaluate a SpEL expression and return the raw value, without locale formatting.
     *
     * <p>Used by template variable declarations ({@code #{let name = expression}}), which must keep
     * the original type so that later expressions can still format it — a number has to stay a
     * number for {@code #fmt.number(...)} to work on the variable.
     *
     * @param expression the expression to evaluate
     * @param rawContext  the context map
     * @param locale     optional locale, or null
     * @param variables  extra variables exposed as {@code #name}; may be {@code null}
     * @return the raw result, or {@link #UNRESOLVED} when evaluation fails
     */
    static Object evaluateRaw(String expression, Map<String, Object> rawContext, Locale locale,
                              Map<String, Object> variables) {
        Result result = evaluateInternal(expression, rawContext, locale, variables);
        return result.resolved() ? result.value() : UNRESOLVED;
    }

    private static Result evaluateInternal(String expression, Map<String, Object> rawContext,
                                           Locale locale, Map<String, Object> variables) {
        Map<String, Object> context = rawContext != null ? rawContext : Collections.emptyMap();

        // A missing SpEL variable evaluates to null instead of failing, which would silently turn a
        // typo into an empty string in a legal document. Reject unknown variables up front so the
        // placeholder stays visible, exactly like a missing input field.
        String unknownVariable = findUnknownVariable(expression, variables);
        if (unknownVariable != null) {
            log.warn("Unknown template variable '#{}' in expression '{}' — placeholder left visible",
                    unknownVariable, expression);
            return Result.failure();
        }

        try {
            EvaluationContext evalContext = SimpleEvaluationContext
                    .forPropertyAccessors(new MapAccessor())
                    .withInstanceMethods()
                    .build();
            if (variables != null) {
                variables.forEach(evalContext::setVariable);
            }
            Expression spelExpression = PARSER.parseExpression(expression);
            return Result.of(spelExpression.getValue(evalContext, context));
        } catch (Exception e) {
            log.warn("SpEL expression evaluation failed for '{}': {}", expression, e.getMessage());
            return Result.failure();
        }
    }

    /**
     * Find the first {@code #variable} in an expression that is neither predefined nor registered.
     *
     * <p>String literals are blanked out first, so {@code '#'} inside a literal is not mistaken for
     * a reference. The captured name stops at the first dot, because {@code #fmt.date(...)} refers
     * to the {@code fmt} variable.
     *
     * @param expression the expression to scan
     * @param variables  the variables available to the expression; may be {@code null}
     * @return the first unknown variable name, or {@code null} when every reference is known
     */
    static String findUnknownVariable(String expression, Map<String, Object> variables) {
        if (expression == null || expression.indexOf('#') < 0) {
            return null;
        }

        String stripped = STRING_LITERAL_PATTERN.matcher(expression).replaceAll(" ");
        Matcher matcher = VARIABLE_REFERENCE_PATTERN.matcher(stripped);

        while (matcher.find()) {
            String name = matcher.group(1);
            if (PREDEFINED_VARIABLES.contains(name)) {
                continue;
            }
            if (variables == null || !variables.containsKey(name)) {
                return name;
            }
        }
        return null;
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
     *   <li>{@code "#salutation"} → {@code []} (a {@code #}-prefixed name is a template-internal
     *       variable, not an input field)</li>
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

            // A '#'-prefixed identifier is a template-internal SpEL variable (#fmt, #salutation),
            // never an input field — skip it so template discovery stays a pure contract view.
            if (matcher.start() > 0 && stripped.charAt(matcher.start() - 1) == '#') {
                continue;
            }

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
