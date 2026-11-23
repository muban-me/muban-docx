package me.muban.docx;

import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
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
 * <p>Both {@code #{else}} and {@code #{elseif ...}} are optional. A block may contain any number
 * of {@code #{elseif}} branches, but {@code #{else}}, when present, must be the last branch:
 *
 * <pre>
 * #{if customer_type == 'CORPORATE'}
 * Corporate wording
 * #{elseif customer_type == 'BIG'}
 * BIG (business information group) wording
 * #{else}
 * Consumer wording
 * #{fi}
 * </pre>
 *
 * <p><b>Nesting:</b> conditional blocks may be nested arbitrarily deep. The inner block must be
 * fully contained in the branch it belongs to:
 *
 * <pre>
 * #{if has_installment_plan}
 *   #{if is_overdue}
 *   Installment plan — payment overdue
 *   #{else}
 *   Installment plan — current
 *   #{fi}
 * #{fi}
 * </pre>
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
 *   <li><b>Paragraph-level only</b> — markers must be standalone paragraphs</li>
 *   <li><b>Balanced</b> — the matcher counts {@code #{if}} depth, so an unterminated inner block
 *       consumes its parent’s {@code #{fi}} and the construct is reported as unmatched</li>
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

    /**
     * Pattern to match {@code #{elseif expression}} — an intermediate branch of a conditional
     * block. {@code #{elif expression}} is accepted as an alias.
     */
    static final Pattern ELSEIF_PATTERN = Pattern.compile("^\\s*#\\{\\s*(?:elseif|elif)\\s+(.+?)\\s*}\\s*$");

    /** Pattern to match {@code #{else}} — the optional else marker */
    static final Pattern ELSE_PATTERN = Pattern.compile("^\\s*#\\{\\s*else\\s*}\\s*$");

    /** Pattern to match {@code #{fi}} — the closing marker */
    static final Pattern FI_PATTERN = Pattern.compile("^\\s*#\\{\\s*fi\\s*}\\s*$");

    /** Opening delimiter shared by every conditional marker. */
    private static final String MARKER_OPEN = "#{";

    /** Closing delimiter shared by every conditional marker. */
    private static final String MARKER_CLOSE = "}";

    private DocxConditionalProcessor() {}

    /**
     * Process conditional blocks in a content list, removing or keeping content
     * based on expression evaluation.
     *
     * <p>Scans the content list for paragraphs matching {@code #{if expr}}, {@code #{elseif expr}},
     * {@code #{else}}, and {@code #{fi}} markers. For each matched group the first branch whose
     * condition holds wins:
     * <ul>
     *   <li>Branches are tested top-down; {@code #{else}}, when present, always matches</li>
     *   <li>The winning branch is kept; every other branch and all marker paragraphs are removed</li>
     *   <li>When no branch matches and there is no {@code #{else}}, the whole block is removed</li>
     * </ul>
     *
     * <p>Blocks may be nested. Once a block is resolved the loop re-scans from the same position,
     * so inner blocks of the kept branch are handled by the same code path.
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
        return processConditionals(content, rawContext, locale, Collections.emptyMap());
    }

    /**
     * Process conditional blocks, with extra variables reachable through the {@code #} prefix.
     *
     * <p>The engine uses this overload to expose its helpers, so a branch condition may format a
     * value: {@code #{if #fmt.date(due_date, 'yyyy') == '2026'}}.
     *
     * @param content    the mutable content list (e.g., body content, cell content)
     * @param rawContext the evaluation context for SpEL expressions
     * @param locale     optional locale for expression evaluation
     * @param variables  extra variables exposed as {@code #name}; may be {@code null}
     * @return the number of conditional blocks processed
     */
    public static int processConditionals(List<Object> content,
                                           Map<String, Object> rawContext,
                                           Locale locale,
                                           Map<String, Object> variables) {
        int blocksProcessed = 0;

        int i = 0;
        while (i < content.size()) {
            Object obj = DocxXmlUtils.unwrap(content.get(i));
            Marker marker = classify(content.get(i));

            if (marker.kind() == MarkerKind.IF) {
                Block block = findBlock(content, i, marker.expression());

                if (block == null) {
                    log.warn("Unmatched #{{if {}}} at content index {} — no closing #{{fi}} found",
                            marker.expression(), i);
                    i++;
                    continue;
                }

                int branch = applyBlock(content, block, rawContext, locale, variables);
                log.debug("Conditional block #{{if {}}} resolved to branch {}",
                        marker.expression(), branch);

                // Do not increment i — the kept branch is re-scanned from the same index,
                // which is what resolves nested conditional blocks.
                blocksProcessed++;
                continue;
            }

            // Check for orphan markers #{{fi}}, #{{else}}, #{{elseif}}
            if (marker.kind() == MarkerKind.FI) {
                log.warn("Orphan #{{fi}} at content index {} — no matching #{{if}} found", i);
            } else if (marker.kind() == MarkerKind.ELSE) {
                log.warn("Orphan #{{else}} at content index {} — no matching #{{if}} found", i);
            } else if (marker.kind() == MarkerKind.ELSEIF) {
                log.warn("Orphan #{{elseif}} at content index {} — no matching #{{if}} found", i);
            }

            // Recurse into nested content (table cells, etc.)
            if (obj instanceof ContentAccessor accessor && !(obj instanceof P)) {
                blocksProcessed += processConditionals(accessor.getContent(), rawContext, locale, variables);
            }

            i++;
        }

        return blocksProcessed;
    }

    // ==================== BLOCK STRUCTURE ====================

    /** Kind of a marker paragraph. */
    private enum MarkerKind { NONE, IF, ELSEIF, ELSE, FI }

    /**
     * A recognised marker paragraph.
     *
     * @param kind       marker kind
     * @param expression the SpEL expression, or {@code null} for {@code #{else}}
     */
    private record Marker(MarkerKind kind, String expression) {
        static final Marker NONE = new Marker(MarkerKind.NONE, null);
    }

    /** One branch of a conditional block; {@code expression == null} means {@code #{else}}. */
    private static final class Branch {
        final int index;
        final String expression;

        Branch(int index, String expression) {
            this.index = index;
            this.expression = expression;
        }
    }

    /** A resolved conditional block: its branches (first one is the {@code #{if}}) and its {@code #{fi}}. */
    private static final class Block {
        final List<Branch> branches;
        /** {@code #{elseif}} markers found after {@code #{else}} — reported once, then dropped. */
        final Set<Integer> ignored;
        final int fiIndex;

        Block(List<Branch> branches, Set<Integer> ignored, int fiIndex) {
            this.branches = branches;
            this.ignored = ignored;
            this.fiIndex = fiIndex;
        }
    }

    /**
     * Classify a content item as a conditional marker.
     *
     * <p>Merges runs split by Word before testing, so a marker broken across runs is still
     * recognised.
     *
     * @param contentItem an element of a content list
     * @return the marker, or {@link Marker#NONE} when the item is not a marker paragraph
     */
    private static Marker classify(Object contentItem) {
        Object obj = DocxXmlUtils.unwrap(contentItem);
        if (!(obj instanceof P paragraph)) {
            return Marker.NONE;
        }

        DocxRunUtils.joinSplitMarker(paragraph, MARKER_OPEN, MARKER_CLOSE);
        String text = DocxXmlUtils.extractAllText(List.of(paragraph)).trim();

        Matcher ifMatcher = IF_PATTERN.matcher(text);
        if (ifMatcher.matches()) {
            return new Marker(MarkerKind.IF, ifMatcher.group(1).trim());
        }

        Matcher elseIfMatcher = ELSEIF_PATTERN.matcher(text);
        if (elseIfMatcher.matches()) {
            return new Marker(MarkerKind.ELSEIF, elseIfMatcher.group(1).trim());
        }

        if (ELSE_PATTERN.matcher(text).matches()) {
            return new Marker(MarkerKind.ELSE, null);
        }

        if (FI_PATTERN.matcher(text).matches()) {
            return new Marker(MarkerKind.FI, null);
        }

        return Marker.NONE;
    }

    /**
     * Resolve the structure of the block opened at {@code ifIndex}.
     *
     * <p>Walks the content list counting {@code #{if}} depth so that nested blocks are skipped
     * rather than mistaken for the closing marker of the current block. Markers at depth 0 are
     * collected as additional branches.
     *
     * @param content      the content list to search
     * @param ifIndex      index of the {@code #{if}} paragraph
     * @param ifExpression expression of the opening {@code #{if}}
     * @return the resolved block, or {@code null} when no closing {@code #{fi}} exists
     */
    private static Block findBlock(List<Object> content, int ifIndex, String ifExpression) {
        List<Branch> branches = new ArrayList<>();
        branches.add(new Branch(ifIndex, ifExpression));

        Set<Integer> ignored = new LinkedHashSet<>();
        boolean sawElse = false;
        int depth = 0;

        for (int j = ifIndex + 1; j < content.size(); j++) {
            Marker marker = classify(content.get(j));

            if (marker.kind() == MarkerKind.IF) {
                depth++;
            } else if (marker.kind() == MarkerKind.FI) {
                if (depth == 0) {
                    return new Block(branches, ignored, j);
                }
                depth--;
            } else if (depth == 0 && marker.kind() == MarkerKind.ELSEIF) {
                if (sawElse) {
                    log.warn("Ignoring #{{elseif}} at content index {} — #{{else}} must be the last branch", j);
                    ignored.add(j);
                } else {
                    branches.add(new Branch(j, marker.expression()));
                }
            } else if (depth == 0 && marker.kind() == MarkerKind.ELSE) {
                if (sawElse) {
                    log.warn("Ignoring duplicate #{{else}} at content index {}", j);
                    ignored.add(j);
                } else {
                    sawElse = true;
                    branches.add(new Branch(j, null));
                }
            }
        }

        return null;
    }

    /**
     * Select the first branch whose condition holds and keep only its body, dropping all marker
     * paragraphs and every other branch.
     *
     * @return the index of the winning branch within {@code block.branches}, or -1 when no branch
     *         matched (the whole block is then removed)
     */
    private static int applyBlock(List<Object> content, Block block,
                                  Map<String, Object> rawContext, Locale locale,
                                  Map<String, Object> variables) {
        int chosen = -1;

        for (int k = 0; k < block.branches.size(); k++) {
            Branch branch = block.branches.get(k);
            if (branch.expression == null) {
                chosen = k;  // #{else} always matches
                break;
            }
            if (evaluateCondition(branch.expression, rawContext, locale, variables)) {
                chosen = k;
                break;
            }
        }

        int firstIndex = block.branches.get(0).index;
        int bodyFrom;
        int bodyTo;

        if (chosen < 0) {
            bodyFrom = firstIndex;
            bodyTo = firstIndex;  // nothing kept
        } else {
            bodyFrom = block.branches.get(chosen).index + 1;
            bodyTo = (chosen + 1 < block.branches.size())
                    ? block.branches.get(chosen + 1).index
                    : block.fiIndex;
        }

        List<Object> kept = new ArrayList<>(content.size());
        kept.addAll(content.subList(0, firstIndex));
        for (int j = bodyFrom; j < bodyTo; j++) {
            if (!block.ignored.contains(j)) {
                kept.add(content.get(j));
            }
        }
        kept.addAll(content.subList(block.fiIndex + 1, content.size()));

        content.clear();
        content.addAll(kept);

        return chosen;
    }

    /**
     * Find the closing {@code #{fi}} marker and the first {@code #{else}} marker at depth 0
     * starting from a given index.
     *
     * <p>Nested {@code #{if}} blocks are skipped by tracking depth, so the returned {@code #{fi}}
     * is the one that closes the block being scanned.
     *
     * @param content   the content list to search
     * @param startFrom index to start searching from (exclusive of the #{if} itself)
     * @return int array: [elseIndex, fiIndex]. Either may be -1 if not found.
     */
    static int[] findClosingMarkers(List<Object> content, int startFrom) {
        int elseIndex = -1;
        int depth = 0;

        for (int j = startFrom; j < content.size(); j++) {
            Marker marker = classify(content.get(j));

            if (marker.kind() == MarkerKind.IF) {
                depth++;
            } else if (marker.kind() == MarkerKind.FI) {
                if (depth == 0) {
                    return new int[]{elseIndex, j};
                }
                depth--;
            } else if (depth == 0 && elseIndex < 0 && marker.kind() == MarkerKind.ELSE) {
                elseIndex = j;
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
        return evaluateCondition(expression, rawContext, locale, Collections.emptyMap());
    }

    /**
     * Evaluate a SpEL expression with extra variables and coerce the result to boolean.
     *
     * @param expression the condition
     * @param rawContext the evaluation context
     * @param locale     optional locale, or null
     * @param variables  extra variables exposed as {@code #name}; may be {@code null}
     * @return the coerced condition value
     */
    static boolean evaluateCondition(String expression, Map<String, Object> rawContext, Locale locale,
                                     Map<String, Object> variables) {
        String result = DocxExpressionEvaluator.evaluate(expression, rawContext, locale, variables);

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

    // ==================== VARIABLE EXTRACTION ====================

    /**
     * Extract variable references from all {@code #{if expr}} conditional markers
     * found in a content tree.
     *
     * <p>Scans paragraphs for {@code #{if ...}} markers (same pattern as runtime processing)
     * and delegates to {@link DocxExpressionEvaluator#extractVariableReferences(String)}
     * to parse variable identifiers from each expression.
     *
     * <p>This method does <b>not</b> modify the content — it is read-only and intended
     * for use during template upload to discover referenced variables.
     *
     * @param content the content list to scan (body, header/footer, table cell, etc.)
     * @return ordered set of variable names referenced in conditional expressions
     */
    public static Set<String> extractConditionalVariables(List<Object> content) {
        Set<String> variables = new LinkedHashSet<>();
        collectConditionalVariables(content, variables);
        return variables;
    }

    /**
     * Recursively collect variable references from conditional markers.
     */
    private static void collectConditionalVariables(List<Object> content, Set<String> variables) {
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);

            if (unwrapped instanceof P paragraph) {
                DocxRunUtils.joinSplitMarker(paragraph, MARKER_OPEN, MARKER_CLOSE);
                String text = DocxXmlUtils.extractAllText(List.of(paragraph)).trim();
                Matcher ifMatcher = IF_PATTERN.matcher(text);
                if (ifMatcher.matches()) {
                    String expression = ifMatcher.group(1).trim();
                    variables.addAll(DocxExpressionEvaluator.extractVariableReferences(expression));
                } else {
                    Matcher elseIfMatcher = ELSEIF_PATTERN.matcher(text);
                    if (elseIfMatcher.matches()) {
                        String expression = elseIfMatcher.group(1).trim();
                        variables.addAll(DocxExpressionEvaluator.extractVariableReferences(expression));
                    }
                }
            }

            if (unwrapped instanceof ContentAccessor accessor && !(unwrapped instanceof P)) {
                collectConditionalVariables(accessor.getContent(), variables);
            }
        }
    }
}
