package me.muban.docx;

import org.docx4j.XmlUtils;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Processes DOCX tables for placeholder replacement and row replication.
 *
 * <p>When a table row contains array-bound placeholders (e.g., {@code ${items.name}}),
 * the row is treated as a template: cloned once per array element, with each clone's
 * placeholders filled from the corresponding data row. The original template row is removed.
 *
 * <p>Non-array rows are processed for scalar placeholder replacement.
 *
 * <p><b>Supported table syntax:</b>
 * <pre>
 * | Header 1       | Header 2    | Header 3       |    ← static header row
 * | ${items.name}  | ${items.qty}| ${items.price}  |   ← template row (cloned per item)
 * </pre>
 */
public class DocxTableProcessor {

    private static final Logger log = LoggerFactory.getLogger(DocxTableProcessor.class);

    /** Pattern to detect array-bound placeholders like ${items.name} — captures array key (group 1) */
    private static final Pattern ARRAY_PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{(\\w+)\\.(\\w+)}");

    /** Pattern to match ${placeholder} expressions */
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    private DocxTableProcessor() {}

    /**
     * Process a table for placeholder replacement and row replication.
     *
     * <p>For each row in the table, checks whether it contains array-bound placeholders
     * (e.g., ${items.name}). If the array key matches a data array, the row is cloned
     * once per array element, and each clone's placeholders are filled with that element's values.
     * The original template row is then removed.
     *
     * <p>Non-array rows are processed normally for scalar placeholder replacement.
     *
     * @param table the docx4j table element
     * @param rawContext raw typed context for SpEL expression evaluation
     * @param dataArrays array data for table row replication (arrayKey → list of row maps)
     * @param locale optional locale for number formatting in array rows, or null
     * @param contentProcessor callback to process content elements (paragraphs, nested tables)
     */
    public static void processTable(Tbl table,
                                     Map<String, Object> rawContext,
                                     Map<String, List<Map<String, Object>>> dataArrays, Locale locale,
                                     ContentProcessor contentProcessor) {
        // Work on a snapshot of the content list to safely modify it during iteration
        List<Object> rows = new ArrayList<>(table.getContent());

        for (Object rowObj : rows) {
            Object unwrappedRow = DocxXmlUtils.unwrap(rowObj);
            if (!(unwrappedRow instanceof Tr tr)) {
                continue;
            }

            // Extract all text from the row to detect array bindings
            String rowText = extractRowText(tr);
            String arrayKey = detectArrayBinding(rowText, dataArrays);

            if (arrayKey != null) {
                // This row is a template row bound to a data array — replicate it
                replicateRow(table, rowObj, tr, arrayKey, dataArrays.get(arrayKey), locale, rawContext);
            } else {
                // Regular row — perform scalar placeholder replacement
                for (Object cell : tr.getContent()) {
                    Object unwrappedCell = DocxXmlUtils.unwrap(cell);
                    if (unwrappedCell instanceof Tc tc) {
                        contentProcessor.processContent(tc.getContent(), rawContext, dataArrays, locale);
                    }
                }
            }
        }
    }

    /**
     * Detect whether a row's text contains array-bound placeholders.
     *
     * <p>An array-bound placeholder has the form ${arrayKey.fieldName} where arrayKey
     * matches a key in the data arrays map. Returns the array key if found, null otherwise.
     */
    static String detectArrayBinding(String rowText, Map<String, List<Map<String, Object>>> dataArrays) {
        if (dataArrays == null || dataArrays.isEmpty()) {
            return null;
        }

        Matcher matcher = ARRAY_PLACEHOLDER_PATTERN.matcher(rowText);
        while (matcher.find()) {
            String candidateKey = matcher.group(1);
            if (dataArrays.containsKey(candidateKey)) {
                return candidateKey;
            }
        }
        return null;
    }

    /**
     * Replicate a template row for each item in a data array.
     *
     * <p>Deep-clones the template row for each array element, replaces ${arrayKey.field}
     * placeholders with the element's field values, inserts the cloned rows at the template
     * row's position, and removes the original template row.
     *
     * <p>Preserves all formatting (borders, shading, fonts, alignment) from the template row.
     */
    private static void replicateRow(Tbl table, Object originalRowObj, Tr templateRow,
                                      String arrayKey, List<Map<String, Object>> items,
                                      Locale locale, Map<String, Object> globalRawContext) {
        int insertIndex = table.getContent().indexOf(originalRowObj);

        if (items == null || items.isEmpty()) {
            log.debug("Data array '{}' is empty, removing template row", arrayKey);
            table.getContent().remove(originalRowObj);
            return;
        }

        log.debug("Replicating template row for array '{}': {} rows", arrayKey, items.size());

        for (Map<String, Object> item : items) {
            // Deep-clone the template row (preserves all XML structure and formatting)
            Tr clonedRow = XmlUtils.deepCopy(templateRow);

            // Build row-scoped raw context for SpEL evaluation:
            // global context + current row item fields accessible by simple name
            // Override the array key with the current item so ${arrayKey.field} navigates correctly
            Map<String, Object> rowRawContext = new LinkedHashMap<>(globalRawContext);
            rowRawContext.putAll(item);
            rowRawContext.put(arrayKey, item);

            // Replace placeholders in each cell of the cloned row
            for (Object cell : clonedRow.getContent()) {
                Object unwrappedCell = DocxXmlUtils.unwrap(cell);
                if (unwrappedCell instanceof Tc tc) {
                    for (Object cellContent : tc.getContent()) {
                        Object unwrappedContent = DocxXmlUtils.unwrap(cellContent);
                        if (unwrappedContent instanceof P paragraph) {
                            replaceParagraphPlaceholders(paragraph, rowRawContext, locale);
                        }
                    }
                }
            }

            // Insert the filled row at the current position
            table.getContent().add(insertIndex++, clonedRow);
        }

        // Remove the original template row
        table.getContent().remove(originalRowObj);

        log.debug("Replicated {} rows for data array '{}'", items.size(), arrayKey);
    }

    /**
     * Replace placeholders in a single paragraph within a replicated row.
     * All placeholders are evaluated via SpEL against the row-scoped raw context.
     */
    private static void replaceParagraphPlaceholders(P paragraph,
                                                      Map<String, Object> rawContext, Locale locale) {
        for (Object obj : paragraph.getContent()) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (!(unwrapped instanceof R run)) continue;

            String text = DocxXmlUtils.getRunText(run);
            if (!text.contains("${")) continue;

            Matcher matcher = PLACEHOLDER_PATTERN.matcher(text);
            StringBuilder result = new StringBuilder();
            while (matcher.find()) {
                String body = matcher.group(1).trim();
                String replacement = DocxExpressionEvaluator.evaluate(body, rawContext, locale);
                matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
            }
            matcher.appendTail(result);
            DocxXmlUtils.setRunText(run, result.toString());
        }
    }

    /**
     * Extract all text from a table row (for array binding detection).
     */
    private static String extractRowText(Tr row) {
        StringBuilder sb = new StringBuilder();
        for (Object cell : row.getContent()) {
            Object unwrappedCell = DocxXmlUtils.unwrap(cell);
            if (unwrappedCell instanceof Tc tc) {
                sb.append(DocxXmlUtils.extractAllText(tc.getContent()));
            }
        }
        return sb.toString();
    }

    /**
     * Callback interface for processing content elements within table cells.
     * Allows the table processor to delegate back to the engine's content processing
     * without creating a circular dependency.
     */
    @FunctionalInterface
    public interface ContentProcessor {
        void processContent(List<Object> content,
                           Map<String, Object> rawContext,
                           Map<String, List<Map<String, Object>>> dataArrays, Locale locale);
    }
}
