package me.muban.docx;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Date;
import java.util.Locale;

/**
 * Formatting helpers exposed to templates under the {@code #fmt} SpEL variable.
 *
 * <p>The SpEL sandbox used by the engine ({@link DocxExpressionEvaluator}) blocks static method
 * calls and constructors, so a template cannot build its own {@link DateTimeFormatter}:
 * {@code T(java.time.format.DateTimeFormatter).ofPattern(...)} fails with {@code EL1005E}.
 * This class closes that gap by injecting a pre-configured object into the evaluation context,
 * so presentation details — date patterns and number patterns — live in the template instead of
 * travelling as pre-formatted strings in the input data.
 *
 * <p><b>Date formatting:</b>
 * <pre>{@code
 * ${#fmt.date(document_date, 'dd.MM.yyyy')}      → 23.11.2026
 * ${#fmt.date(document_date, 'd MMMM yyyy')}     → 23 listopada 2026   (genitive, format-wide)
 * ${#fmt.date(document_date, 'LLLL yyyy')}       → listopad 2026       (nominative, stand-alone)
 * }</pre>
 *
 * <p>The input may be a {@link TemporalAccessor} ({@code LocalDate}, {@code LocalDateTime}, …),
 * a {@link Date}, or a string. ISO-8601 strings are the expected contract; a few common
 * non-ISO shapes ({@code dd.MM.yyyy}, {@code dd-MM-yyyy}, {@code dd/MM/yyyy},
 * {@code yyyy-MM-dd HH:mm[:ss]}) are accepted so a template does not break on legacy data.
 * A value that cannot be parsed is returned unchanged and logged at {@code warn} level, so the
 * problem is visible in the output instead of silently producing an empty cell.
 *
 * <p><b>Number formatting:</b>
 * <pre>{@code
 * ${#fmt.number(debt_total_amount, '#,##0.00')}  → 1 234,50   (pl_PL)
 * }</pre>
 *
 * <p><b>Scope.</b> This helper is deliberately domain-neutral: it knows patterns and locales, not
 * business vocabulary. Anything that needs a dictionary — amounts in words, month names in the
 * locative case, declension of nouns or forms of address — stays out of the library. A template
 * that needs a wording that varies with a data value expresses it as an ordinary expression, which
 * also keeps the text visible next to its context:
 * <pre>{@code
 * ${addressee_gender == 'F' ? 'Dear Ms,' : 'Dear Mr,'}
 * }</pre>
 */
public final class DocxFormatHelper {

    private static final Logger log = LoggerFactory.getLogger(DocxFormatHelper.class);

    /** Date pattern used by {@link #date(Object)} when no explicit pattern is given. */
    public static final String DEFAULT_DATE_PATTERN = "dd.MM.yyyy";

    /** Number pattern used by {@link #number(Object)} when no explicit pattern is given. */
    public static final String DEFAULT_NUMBER_PATTERN = "#,##0.00";

    /** Date-only fallback patterns accepted besides ISO-8601. */
    private static final DateTimeFormatter[] FALLBACK_DATE_PATTERNS = {
            DateTimeFormatter.ofPattern("dd.MM.yyyy"),
            DateTimeFormatter.ofPattern("dd-MM-yyyy"),
            DateTimeFormatter.ofPattern("dd/MM/yyyy")
    };

    /** Date-time fallback patterns accepted besides ISO-8601. */
    private static final DateTimeFormatter[] FALLBACK_DATE_TIME_PATTERNS = {
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
            DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm")
    };

    private final Locale locale;

    /** Creates a helper bound to the JVM default locale. */
    public DocxFormatHelper() {
        this(Locale.getDefault());
    }

    /**
     * Creates a helper bound to the given locale.
     *
     * <p>The locale controls month names, month forms, grouping and decimal separators; the
     * pattern controls the shape. A {@code null} locale falls back to the JVM default.
     *
     * @param locale the document locale, or {@code null}
     */
    public DocxFormatHelper(Locale locale) {
        this.locale = locale != null ? locale : Locale.getDefault();
    }

    /**
     * The locale this helper renders with.
     *
     * @return the effective locale, never {@code null}
     */
    public Locale getLocale() {
        return locale;
    }

    /**
     * Format a date using an explicit {@link DateTimeFormatter} pattern.
     *
     * @param value   the date to format: {@link TemporalAccessor}, {@link Date}, or string
     * @param pattern the pattern, for example {@code dd.MM.yyyy} or {@code d MMMM yyyy}
     * @return the formatted date; {@code null} for {@code null} input; the original value when it
     *         cannot be parsed or the pattern cannot be applied to it
     */
    public String date(Object value, String pattern) {
        if (value == null) {
            return null;
        }
        String effectivePattern = pattern == null || pattern.isBlank() ? DEFAULT_DATE_PATTERN : pattern.trim();

        try {
            if (value instanceof Date date) {
                // Format the legacy type directly: converting through Instant would shift the
                // calendar day for values parsed in the local zone.
                return new SimpleDateFormat(effectivePattern, locale).format(date);
            }

            TemporalAccessor temporal = toTemporal(value);
            if (temporal == null) {
                log.warn("Cannot parse '{}' as a date — value left unchanged in the output", value);
                return String.valueOf(value);
            }
            return DateTimeFormatter.ofPattern(effectivePattern, locale).format(temporal);
        } catch (RuntimeException e) {
            log.warn("Cannot format date '{}' with pattern '{}': {}", value, effectivePattern, e.getMessage());
            return String.valueOf(value);
        }
    }

    /**
     * Format a date using {@link #DEFAULT_DATE_PATTERN}.
     *
     * @param value the date to format
     * @return the formatted date, or the original value when it cannot be parsed
     */
    public String date(Object value) {
        return date(value, DEFAULT_DATE_PATTERN);
    }

    /**
     * Format a number using an explicit {@link DecimalFormat} pattern.
     *
     * <p>The pattern always uses {@code .} as the decimal separator and {@code ,} as the grouping
     * separator (Java convention); the locale decides how those appear in the result. Floating
     * point values are converted through their string form, so money amounts do not pick up
     * binary-representation artifacts.
     *
     * @param value   the number to format: {@link Number} or a numeric string
     * @param pattern the pattern, for example {@code #,##0.00}
     * @return the formatted number; {@code null} for {@code null} input; the original value when
     *         it is not numeric or the pattern is invalid
     */
    public String number(Object value, String pattern) {
        if (value == null) {
            return null;
        }
        String effectivePattern = pattern == null || pattern.isBlank() ? DEFAULT_NUMBER_PATTERN : pattern.trim();

        BigDecimal number = toBigDecimal(value);
        if (number == null) {
            log.warn("Cannot parse '{}' as a number — value left unchanged in the output", value);
            return String.valueOf(value);
        }

        try {
            DecimalFormat format = new DecimalFormat(
                    effectivePattern, DecimalFormatSymbols.getInstance(locale));
            return format.format(number);
        } catch (IllegalArgumentException e) {
            log.warn("Cannot format number '{}' with pattern '{}': {}", value, effectivePattern, e.getMessage());
            return String.valueOf(value);
        }
    }

    /**
     * Format a number using {@link #DEFAULT_NUMBER_PATTERN}.
     *
     * @param value the number to format
     * @return the formatted number, or the original value when it is not numeric
     */
    public String number(Object value) {
        return number(value, DEFAULT_NUMBER_PATTERN);
    }

    // ==================== CONVERSION ====================

    private static TemporalAccessor toTemporal(Object value) {
        if (value instanceof TemporalAccessor temporal) {
            return temporal;
        }
        if (!(value instanceof CharSequence sequence)) {
            return null;
        }
        String text = sequence.toString().trim();
        if (text.isEmpty()) {
            return null;
        }

        TemporalAccessor parsed = parseIsoDateTime(text);
        if (parsed != null) {
            return parsed;
        }
        parsed = parseIsoDate(text);
        if (parsed != null) {
            return parsed;
        }
        for (DateTimeFormatter formatter : FALLBACK_DATE_TIME_PATTERNS) {
            try {
                return LocalDateTime.parse(text, formatter);
            } catch (RuntimeException ignored) {
                // try the next pattern
            }
        }
        for (DateTimeFormatter formatter : FALLBACK_DATE_PATTERNS) {
            try {
                return LocalDate.parse(text, formatter);
            } catch (RuntimeException ignored) {
                // try the next pattern
            }
        }
        return null;
    }

    private static TemporalAccessor parseIsoDateTime(String text) {
        try {
            return OffsetDateTime.parse(text, DateTimeFormatter.ISO_OFFSET_DATE_TIME);
        } catch (RuntimeException ignored) {
            // not an offset date-time
        }
        try {
            return LocalDateTime.parse(text, DateTimeFormatter.ISO_LOCAL_DATE_TIME);
        } catch (RuntimeException ignored) {
            return null;
        }
    }

    private static TemporalAccessor parseIsoDate(String text) {
        try {
            return LocalDate.parse(text, DateTimeFormatter.ISO_LOCAL_DATE);
        } catch (RuntimeException ignored) {
            return null;
        }
    }

    private static BigDecimal toBigDecimal(Object value) {
        if (value instanceof BigDecimal decimal) {
            return decimal;
        }
        if (value instanceof Number number) {
            // Via the string form: binary floating point would otherwise leak into money amounts.
            return new BigDecimal(number.toString());
        }
        if (value instanceof CharSequence sequence) {
            String text = sequence.toString()
                    .replace("\u00A0", "")   // no-break space, used as a grouping separator
                    .replace("\u202F", "")   // narrow no-break space
                    .replace(" ", "")
                    .trim();
            if (text.isEmpty()) {
                return null;
            }
            try {
                return new BigDecimal(text.replace(',', '.'));
            } catch (NumberFormatException e) {
                return null;
            }
        }
        return null;
    }
}
