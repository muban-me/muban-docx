package me.muban.docx;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.NumberFormat;
import java.util.Locale;

/**
 * Locale utilities for parsing locale strings and formatting numeric values.
 */
public final class LocaleUtils {

    private static final Logger log = LoggerFactory.getLogger(LocaleUtils.class);

    private LocaleUtils() {}

    /**
     * Parse a locale string (e.g. {@code "pl_PL"}, {@code "en"}) into a {@link Locale}.
     *
     * @param localeString locale in {@code language}, {@code language_country},
     *                     or {@code language_country_variant} format
     * @return the parsed Locale
     * @throws IllegalArgumentException if the string is null, empty, or malformed
     */
    public static Locale parseLocale(String localeString) {
        if (localeString == null || localeString.trim().isEmpty()) {
            throw new IllegalArgumentException("Locale string cannot be null or empty");
        }
        String[] parts = localeString.trim().split("_");
        if (parts.length == 1) {
            return new Locale(parts[0]);
        } else if (parts.length == 2) {
            return new Locale(parts[0], parts[1]);
        } else if (parts.length == 3) {
            return new Locale(parts[0], parts[1], parts[2]);
        } else {
            throw new IllegalArgumentException(
                    "Invalid locale format: " + localeString +
                    ". Expected format: 'language' or 'language_country' or 'language_country_variant'");
        }
    }

    /**
     * Format a value for display. Numbers are formatted using the given locale
     * (grouping separators, decimal separator). Other types are converted via {@code String.valueOf()}.
     *
     * @param value  the value to format (may be null)
     * @param locale the locale for number formatting, or null for no locale-aware formatting
     * @return formatted string, or null if value is null
     */
    public static String formatValue(Object value, Locale locale) {
        if (value == null) {
            return null;
        }
        if (value instanceof Number number && locale != null) {
            String formatted = NumberFormat.getInstance(locale).format(number);
            log.trace("Formatted number {} → '{}' using locale {}", value, formatted, locale);
            return formatted;
        }
        return String.valueOf(value);
    }
}
