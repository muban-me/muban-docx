package me.muban.docx;

import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Processes template variable declarations written as {@code #{let name = expression}}.
 *
 * <p>This is the DOCX counterpart of a JasperReports report variable: a value computed once in the
 * template and reused wherever it is needed, instead of repeating the same (often lengthy)
 * expression in every placeholder. Declarations live in <b>their own paragraphs</b>, like the
 * conditional markers they sit next to, and are removed from the output.
 *
 * <pre>
 * #{let salutation = "M".equals(addressee_gender) ? 'Szanowny Panie,' : ("F".equals(addressee_gender) ? 'Szanowna Pani,' : 'Szanowni Państwo,')}
 * #{let debtor_nom = addressee_gender == 'F' ? 'dłużniczka' : 'dłużnik'}
 * #{let due_date_long = #fmt.date(debt_payment_date, 'd MMMM yyyy')}
 * </pre>
 *
 * <p>Declared names become ordinary context entries, so they are used as {@code ${salutation}}
 * and may also appear inside conditional expressions ({@code #{if debtor_nom == ...}}).
 *
 * <h3>Semantics</h3>
 * <ul>
 *   <li><b>Sequential</b> — declarations are evaluated in document order; a later declaration sees
 *       the ones before it. A forward reference simply fails to evaluate, so no cycle detection is
 *       needed.</li>
 *   <li><b>Whole document</b> — the body and headers/footers share one namespace.</li>
 *   <li><b>Raw value</b> — the evaluated object is stored as-is (a number stays a number), so
 *       {@code #fmt.number(...)} can still format it; locale formatting happens at output time.</li>
 *   <li><b>Overrides are visible</b> — a declaration that shadows an input field is accepted but
 *       logged as a warning, because silent shadowing hides data problems.</li>
 *   <li><b>Failure is loud</b> — an expression that cannot be evaluated leaves the variable
 *       undefined, so {@code ${name}} stays visible in the output instead of rendering as blank.</li>
 *   <li><b>Document scope only</b> — declarations must not be placed inside a replicated table row:
 *       this phase runs before row replication, so such a declaration would be evaluated once.</li>
 * </ul>
 *
 * <p>Because declarations are resolved before conditional blocks, a {@code #{let}} cannot be
 * assigned conditionally — a declaration placed inside an {@code #{if}} block would run
 * unconditionally. Use a ternary inside the declaration instead.
 *
 * @see DocxExpressionEvaluator
 * @see DocxConditionalProcessor
 */
final class DocxVariableProcessor {

    private static final Logger log = LoggerFactory.getLogger(DocxVariableProcessor.class);

    /**
     * Pattern matching a declaration occupying a whole paragraph.
     *
     * <p>The expression capture is greedy so the last {@code }} in the paragraph delimits the
     * marker; that way an expression may itself contain braces (for example an inline map literal).
     * Group 1 is the variable name, group 2 the expression.
     */
    static final Pattern LET_PATTERN = Pattern.compile(
            "^\\s*#\\{\\s*let\\s+([A-Za-z_][A-Za-z0-9_]*)\\s*=\\s*(.+)\\}\\s*$", Pattern.DOTALL);

    /** Cheap pre-check so ordinary paragraphs never reach the full pattern. */
    private static final Pattern LET_PREFIX = Pattern.compile("^\\s*#\\{\\s*let\\s");

    private static final String MARKER_OPEN = "#{";

    private static final String MARKER_CLOSE = "}";

    private DocxVariableProcessor() {}

    /** A recognised {@code #{let name = expression}} declaration. */
    private record Declaration(String name, String expression) {}

    /**
     * Resolve every declaration in a content list, adding the results to {@code context} and
     * removing the marker paragraphs.
     *
     * @param content         the mutable content list (body, header, footer, table cell, …)
     * @param context         the mutable evaluation context; declared names are written into it
     * @param helperVariables variables exposed as {@code #name} (the format helper); may be null
     * @param locale          locale used when evaluating expressions, or null
     * @return the number of declarations successfully resolved
     */
    static int processVariables(List<Object> content,
                                Map<String, Object> context,
                                Map<String, Object> helperVariables,
                                Locale locale) {
        int resolved = 0;
        int i = 0;

        while (i < content.size()) {
            Object obj = DocxXmlUtils.unwrap(content.get(i));

            Declaration declaration = classify(content.get(i));
            if (declaration != null) {
                Object value = DocxExpressionEvaluator.evaluateRaw(
                        declaration.expression(), context, locale, helperVariables);

                // The marker never survives into the output, resolved or not.
                content.remove(i);

                if (value == DocxExpressionEvaluator.UNRESOLVED) {
                    log.warn("Could not evaluate #{{let {} = {}}} — '{}' stays undefined",
                            declaration.name(), declaration.expression(), declaration.name());
                    // Do not advance: the next element slid into this index.
                    continue;
                }

                if (context.containsKey(declaration.name())) {
                    log.warn("Template variable '{}' overrides a value of the same name from the input data",
                            declaration.name());
                }
                if (helperVariables != null && helperVariables.containsKey(declaration.name())) {
                    log.warn("Template variable '{}' shadows a built-in '#' helper", declaration.name());
                }

                context.put(declaration.name(), value);
                log.debug("Template variable '{}' = {}", declaration.name(), value);
                resolved++;
                continue;
            }

            // Recurse into nested content (table cells, content controls, …). Paragraphs are
            // leaves for this phase.
            if (obj instanceof ContentAccessor accessor && !(obj instanceof P)) {
                resolved += processVariables(accessor.getContent(), context, helperVariables, locale);
            }

            i++;
        }

        return resolved;
    }

    /**
     * Collect declarations without evaluating them, in document order.
     *
     * <p>Backs {@link MubanDocxEngine#extractDeclaredVariables}, which lets a caller derive the
     * template's variable list — and with it the set of input fields it depends on.
     *
     * @param content the content list to scan
     * @param sink    ordered map receiving name → expression
     */
    static void collectDeclarations(List<Object> content, Map<String, String> sink) {
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            Declaration declaration = classify(unwrapped);
            if (declaration != null) {
                String previous = sink.putIfAbsent(declaration.name(), declaration.expression());
                if (previous != null) {
                    log.debug("Variable '{}' is declared more than once; the last declaration wins at runtime",
                            declaration.name());
                }
                continue;
            }
            if (unwrapped instanceof ContentAccessor accessor && !(unwrapped instanceof P)) {
                collectDeclarations(accessor.getContent(), sink);
            }
        }
    }

    /**
     * Build an empty ordered sink for {@link #collectDeclarations}.
     *
     * @return a new insertion-ordered map
     */
    static Map<String, String> newDeclarationMap() {
        return new LinkedHashMap<>();
    }

    /**
     * Recognise a declaration paragraph.
     *
     * <p>Merges runs split by Word before matching, exactly like the conditional markers, so a
     * declaration broken across runs is still recognised.
     *
     * @param contentItem an element of a content list
     * @return the declaration, or {@code null} when the item is not one
     */
    private static Declaration classify(Object contentItem) {
        Object obj = DocxXmlUtils.unwrap(contentItem);
        if (!(obj instanceof P paragraph)) {
            return null;
        }

        DocxRunUtils.joinSplitMarker(paragraph, MARKER_OPEN, MARKER_CLOSE);
        String text = DocxXmlUtils.extractAllText(List.of(paragraph)).trim();

        if (!LET_PREFIX.matcher(text).find()) {
            return null;
        }

        Matcher matcher = LET_PATTERN.matcher(text);
        if (!matcher.matches()) {
            log.warn("Malformed #{{let}} marker, ignoring: {}", text);
            return null;
        }
        return new Declaration(matcher.group(1), matcher.group(2).trim());
    }
}
