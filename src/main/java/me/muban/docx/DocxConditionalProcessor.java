package me.muban.docx;

import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Processes conditional blocks in DOCX templates using {@code #{if expr}} / {@code #{else}} /
 * {@code #{fi}} markers.
 *
 * <p>Conditional blocks allow template designers to include or exclude entire sections of
 * a document based on runtime data values. The markers must be placed in <b>their own
 * paragraphs</b> (dedicated lines in the Word document).
 *
 * <p><b>Syntax:</b></p>
 * <pre>
 * #{if customer_debt &gt; 1000}
 * This paragraph will only appear when customer_debt exceeds 1000.
 * #{else}
 * This paragraph appears when debt is 1000 or less.
 * #{fi}
 * </pre>
 *
 * <p>The {@code #{else}} marker is optional.
 *
 * <p><b>Expression evaluation:</b></p>
 * <p>The expression inside {@code #{if ...}} is evaluated via {@link DocxExpressionEvaluator}
 * with the same SpEL security sandbox as {@code ${...}} placeholders. The result is coerced
 * to a boolean:
 * <ul>
 *   <li>{@code Boolean} — used directly</li>
 *   <li>{@code Number} — {@code != 0} is true</li>
 *   <li>{@code String} — non-empty and not {@code "false"} is true</li>
 *   <li>{@code null} — false</li>
 *   <li>Anything else — true</li>
 * </ul>
 *
 * <h3>Design constraints</h3>
 * <ul>
 *   <li><b>No nesting</b> — conditional blocks cannot contain other conditional blocks</li>
 *   <li><b>Paragraph-level only</b> — markers must be standalone paragraphs</li>
 *   <li><b>Runs before substitution</b> — conditional processing happens before
 *       placeholder replacement to avoid evaluating placeholders inside excluded blocks</li>
 * </ul>
 *
 * @see DocxExpressionEvaluator
 */
public final class DocxConditionalProcessor {

    private static final Logger log = LoggerFactory.getLogger(DocxConditionalProcessor.class);

    /** Pattern to match {@code #{if expression}} — the full paragraph text */
    static final Pattern IF_PATTERN = Pattern.compile("^\\s*#\\{\\s*if\\s+(.+?)\\s*}\\s*$");

    /** Pattern to match {@code #{else}} — the optional else marker */
    static final Pattern ELSE_PATTERN = Pattern.compile("^\\s*#\\{\\s*else\\s*}\\s*$");

    /** Pattern to match {@code #{fi}} — the closing marker */
    static final Pattern FI_PATTERN = Pattern.compile("^\\s*#\\{\\s*fi\\s*}\\s*$");

    private DocxConditionalProcessor() {}

    /**
     * Process conditional blocks in a content list, removing or keeping content
     * based on expression evaluation.
     *
     * <p>Scans the content list for paragraphs matching {@code #{if expr}},
     * {@code #{else}}, and {@code #{fi}} markers. For each matched group:
     * <ul>
     *   <li>If the expression evaluates to <b>true</b>: keeps content between {@code #{if}} and
     *       {@code #{else}} (or {@code #{fi}} if no else), removes the rest</li>
     *   <li>If the expression evaluates to <b>false</b>: keeps content between {@code #{else}} and
     *       {@code #{fi}} (if else exists), removes the rest</li>
     * </ul>
     *
     * <p>Unmatched markers are logged as warnings and left in place.
     *
     * @param content    the mutable content list (e.g., body content, cell content)
     * @param rawContext the evaluation context for SpEL expressions
     * @param locale     optional locale for expression evaluation
     * @return the number of conditional blocks processed
     */
    public static int processConditionals(List<Object> content,
                                           Map<String, Object> rawContext,
                                           Locale locale) {
        int blocksProcessed = 0;

        int i = 0;
        while (i < content.size()) {
            Object obj = DocxXmlUtils.unwrap(content.get(i));

            if (obj instanceof P paragraph) {
                joinConditionalRuns(paragraph);

                String text = DocxXmlUtils.extractAllText(List.of(paragraph)).trim();
                Matcher ifMatcher = IF_PATTERN.matcher(text);

                if (ifMatcher.matches()) {
                    String expression = ifMatcher.group(1).trim();

                    int[] markers = findClosingMarkers(content, i + 1);
                    int elseIndex = markers[0];
                    int fiIndex = markers[1];

                    if (fiIndex < 0) {
                        log.warn("Unmatched #{{if {}}} at content index {} — no closing #{{fi}} found", expression, i);
                        i++;
                        continue;
                    }

                    boolean result = evaluateCondition(expression, rawContext, locale);
                    log.debug("Conditional block #{{if {}}} evaluated to {}", expression, result);

                    if (elseIndex < 0) {
                        // No #{else} — simple if/fi block
                        if (result) {
                            content.remove(fiIndex);
                            content.remove(i);
                        } else {
                            content.subList(i, fiIndex + 1).clear();
                        }
                    } else {
                        // Has #{else} — if/else/fi block
                        if (result) {
                            content.subList(elseIndex, fiIndex + 1).clear();
                            content.remove(i);
                        } else {
                            content.remove(fiIndex);
                            content.subList(i, elseIndex + 1).clear();
                        }
                    }
                    // Don't increment i — content is now at index i

                    blocksProcessed++;
                    continue;
                }

                // Check for orphan #{fi} or #{else}
                if (FI_PATTERN.matcher(text).matches()) {
                    log.warn("Orphan #{{fi}} at content index {} — no matching #{{if}} found", i);
                } else if (ELSE_PATTERN.matcher(text).matches()) {
                    log.warn("Orphan #{{else}} at content index {} — no matching #{{if}} found", i);
                }
            }

            // Recurse into nested content (table cells, etc.)
            if (obj instanceof ContentAccessor accessor && !(obj instanceof P)) {
                blocksProcessed += processConditionals(accessor.getContent(), rawContext, locale);
            }

            i++;
        }

        return blocksProcessed;
    }

    /**
     * Find the closing {@code #{fi}} marker and optional {@code #{else}} marker
     * starting from a given index.
     *
     * @param content   the content list to search
     * @param startFrom index to start searching from (exclusive of the #{if} itself)
     * @return int array: [elseIndex, fiIndex]. Either may be -1 if not found.
     */
    static int[] findClosingMarkers(List<Object> content, int startFrom) {
        int elseIndex = -1;
        for (int j = startFrom; j < content.size(); j++) {
            Object obj = DocxXmlUtils.unwrap(content.get(j));
            if (obj instanceof P paragraph) {
                joinConditionalRuns(paragraph);
                String text = DocxXmlUtils.extractAllText(List.of(paragraph)).trim();
                if (FI_PATTERN.matcher(text).matches()) {
                    return new int[]{elseIndex, j};
                }
                if (elseIndex < 0 && ELSE_PATTERN.matcher(text).matches()) {
                    elseIndex = j;
                }
            }
        }
        return new int[]{elseIndex, -1};
    }

    /**
     * Find the closing {@code #{fi}} marker starting from a given index.
     * Convenience method that ignores {@code #{else}}.
     *
     * @param content   the content list to search
     * @param startFrom index to start searching from
     * @return the index of the closing #{fi} paragraph, or -1 if not found
     */
    static int findClosingFi(List<Object> content, int startFrom) {
        return findClosingMarkers(content, startFrom)[1];
    }

    /**
     * Evaluate a SpEL expression and coerce the result to boolean.
     */
    static boolean evaluateCondition(String expression, Map<String, Object> rawContext, Locale locale) {
        String result = DocxExpressionEvaluator.evaluate(expression, rawContext, locale);

        if (result.startsWith("${") && result.endsWith("}")) {
            log.warn("Conditional expression '{}' evaluation failed, treating as false", expression);
            return false;
        }

        return coerceToBoolean(result);
    }

    /**
     * Coerce a string result to boolean.
     *
     * <ul>
     *   <li>null or empty → false</li>
     *   <li>"true" (case-insensitive) → true</li>
     *   <li>"false" (case-insensitive) → false</li>
     *   <li>Numeric string: 0 → false, non-zero → true</li>
     *   <li>Any other non-empty string → true</li>
     * </ul>
     */
    static boolean coerceToBoolean(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }

        if ("true".equalsIgnoreCase(value)) {
            return true;
        }
        if ("false".equalsIgnoreCase(value)) {
            return false;
        }

        try {
            double num = Double.parseDouble(value);
            return num != 0.0;
        } catch (NumberFormatException ignored) {
            // Not a number — any non-empty string is truthy
        }

        return true;
    }

    /**
     * Merge adjacent runs in a paragraph that together form a {@code #{...}} marker.
     *
     * <p>Word often splits {@code #{if expr}} across multiple runs due to spell-checking,
     * proofing marks, or editing history. This method merges adjacent runs with identical
     * formatting when they collectively contain {@code #} and {@code {}.
     */
    static void joinConditionalRuns(P paragraph) {
        List<Object> content = paragraph.getContent();

        int i = 0;
        while (i < content.size()) {
            Object unwrapped = DocxXmlUtils.unwrap(content.get(i));
            if (!(unwrapped instanceof R run)) {
                i++;
                continue;
            }

            String text = DocxXmlUtils.getRunText(run);
            if (!text.contains("#") && !text.contains("{")) {
                i++;
                continue;
            }

            StringBuilder accumulated = new StringBuilder(text);
            int mergeEnd = i;

            for (int j = i + 1; j < content.size(); j++) {
                Object nextUnwrapped = DocxXmlUtils.unwrap(content.get(j));
                if (!(nextUnwrapped instanceof R nextRun)) break;

                if (!isCompatibleFormatting(run, nextRun)) break;

                String nextText = DocxXmlUtils.getRunText(nextRun);
                accumulated.append(nextText);
                mergeEnd = j;

                String acc = accumulated.toString();
                if (acc.contains("#{") && acc.contains("}")) {
                    DocxXmlUtils.setRunText(run, acc);
                    for (int k = mergeEnd; k > i; k--) {
                        content.remove(k);
                    }
                    break;
                }
            }

            i++;
        }
    }

    /**
     * Check if two runs have compatible formatting for merging.
     * Compares serialized RPr XML for deep equality.
     */
    private static boolean isCompatibleFormatting(R run1, R run2) {
        var rpr1 = run1.getRPr();
        var rpr2 = run2.getRPr();

        if (rpr1 == null && rpr2 == null) {
            return true;
        }
        if (rpr1 == null || rpr2 == null) {
            return false;
        }

        try {
            String xml1 = org.docx4j.XmlUtils.marshaltoString(rpr1);
            String xml2 = org.docx4j.XmlUtils.marshaltoString(rpr2);
            return xml1.equals(xml2);
        } catch (Exception e) {
            log.debug("Could not compare RPr for run merging: {}", e.getMessage());
            return false;
        }
    }
}
