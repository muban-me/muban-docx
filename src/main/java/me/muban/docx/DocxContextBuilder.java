package me.muban.docx;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Builds the evaluation context for DOCX template placeholder replacement.
 *
 * <p>Combines parameter values and data map entries into a typed context map
 * where all values retain their original types (numbers, booleans, lists, nested maps).
 * Every {@code ${placeholder}} is evaluated via SpEL against this context,
 * with locale-aware formatting applied at output time by {@link DocxExpressionEvaluator}.
 *
 * <p>Also extracts data arrays (List of Maps) from the data map for table row replication.
 */
public class DocxContextBuilder {

    private static final Logger log = LoggerFactory.getLogger(DocxContextBuilder.class);

    private DocxContextBuilder() {}

    /**
     * Build the evaluation context preserving original value types.
     *
     * <p>All placeholders — both simple keys ({@code ${recipientName}}) and SpEL
     * expressions ({@code ${count > 1 ? "items" : "item"}}) — are evaluated via
     * SpEL against this context. Locale-aware formatting is applied at output time
     * by {@link DocxExpressionEvaluator#evaluate}.
     *
     * <p>Parameters are added first, then data entries overlay (data wins on key conflict).
     * Data entries are NOT flattened — nested maps and arrays remain intact for SpEL
     * navigation (e.g., {@code address.city}, {@code items.size()}).
     *
     * @param parameters simple key-value parameters (all strings); may be null
     * @param data       raw data map (may contain nested maps and arrays); may be null
     * @return unmodifiable map of name → typed value for expression evaluation
     */
    public static Map<String, Object> buildRawContext(Map<String, String> parameters,
                                                       Map<String, Object> data) {
        Map<String, Object> raw = new LinkedHashMap<>();

        // Add parameters (string values)
        if (parameters != null) {
            for (Map.Entry<String, String> entry : parameters.entrySet()) {
                if (entry.getKey() != null && entry.getValue() != null) {
                    raw.put(entry.getKey(), entry.getValue());
                }
            }
        }

        // Add data entries (preserves nested maps and arrays for expression access)
        if (data != null) {
            raw.putAll(data);
        }

        log.debug("Built raw expression context with {} entries: {}", raw.size(), raw.keySet());
        return Collections.unmodifiableMap(raw);
    }

    /**
     * Build the evaluation context from a data-only map (no separate parameters).
     *
     * <p>Convenience overload for when all values come from a single data map.
     *
     * @param data raw data map (may contain nested maps and arrays); may be null
     * @return unmodifiable map of name → typed value for expression evaluation
     */
    public static Map<String, Object> buildRawContext(Map<String, Object> data) {
        if (data == null || data.isEmpty()) {
            return Collections.emptyMap();
        }
        return Collections.unmodifiableMap(new LinkedHashMap<>(data));
    }

    /**
     * Extract List entries from the data map for table row replication.
     *
     * <p>Scans the top-level data map for entries where the value is a List of Maps.
     * Each such entry represents a data array that can drive table row cloning.
     *
     * <p>Example input:
     * <pre>{@code
     * {
     *   "companyName": "ACME Corp",     // scalar — goes to context
     *   "items": [                        // array — goes to dataArrays
     *     {"name": "Widget A", "price": 29.99},
     *     {"name": "Widget B", "price": 14.99}
     *   ]
     * }
     * }</pre>
     *
     * @param data the raw data map from the request
     * @return map of array key → list of row maps (empty map if no arrays found)
     */
    public static Map<String, List<Map<String, Object>>> extractDataArrays(Map<String, Object> data) {
        Map<String, List<Map<String, Object>>> arrays = new LinkedHashMap<>();

        if (data == null) {
            return arrays;
        }

        for (Map.Entry<String, Object> entry : data.entrySet()) {
            if (entry.getValue() instanceof List<?> list && !list.isEmpty()) {
                // Verify the list contains Maps (rows of key-value data)
                if (list.get(0) instanceof Map) {
                    @SuppressWarnings("unchecked")
                    List<Map<String, Object>> typedList = (List<Map<String, Object>>) list;
                    arrays.put(entry.getKey(), typedList);
                    log.debug("Found data array '{}' with {} rows for table replication",
                            entry.getKey(), typedList.size());
                }
            }
        }

        return arrays;
    }
}
