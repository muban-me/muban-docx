package me.muban.docx;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.Locale;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for {@link DocxFormatHelper} — the {@code #fmt} helper injected into the SpEL context.
 *
 * <p>The point of these tests is that presentation patterns live in the template: the helper must
 * turn a raw contract value (usually an ISO string) into the shape the pattern asks for, using the
 * document locale for month names and separators.
 */
@DisplayName("DocxFormatHelper")
class DocxFormatHelperTest {

    private static final Locale PL = Locale.forLanguageTag("pl-PL");

    private final DocxFormatHelper format = new DocxFormatHelper(PL);

    // ==================== DATE ====================

    @Nested
    @DisplayName("date()")
    class DateTests {

        @Test
        @DisplayName("numeric pattern")
        void numericPattern() {
            assertThat(format.date(LocalDate.of(2026, 11, 23), "dd.MM.yyyy"))
                    .isEqualTo("23.11.2026");
        }

        @Test
        @DisplayName("MMMM renders the genitive month form used inside a date")
        void genitiveMonth() {
            assertThat(format.date(LocalDate.of(2026, 11, 23), "d MMMM yyyy"))
                    .isEqualTo("23 listopada 2026");
        }

        @Test
        @DisplayName("LLLL renders the stand-alone nominative month form")
        void nominativeMonth() {
            assertThat(format.date(LocalDate.of(2026, 11, 23), "LLLL yyyy"))
                    .isEqualTo("listopad 2026");
        }

        @Test
        @DisplayName("ISO date string is accepted")
        void isoDateString() {
            assertThat(format.date("2026-11-23", "dd.MM.yyyy"))
                    .isEqualTo("23.11.2026");
        }

        @Test
        @DisplayName("ISO date-time string is accepted")
        void isoDateTimeString() {
            assertThat(format.date("2026-11-23T10:15:30", "dd.MM.yyyy HH:mm"))
                    .isEqualTo("23.11.2026 10:15");
        }

        @Test
        @DisplayName("LocalDateTime keeps the time part")
        void localDateTime() {
            assertThat(format.date(LocalDateTime.of(2026, 11, 23, 10, 15), "dd.MM.yyyy HH:mm"))
                    .isEqualTo("23.11.2026 10:15");
        }

        @Test
        @DisplayName("legacy dd.MM.yyyy input is accepted")
        void legacyInput() {
            assertThat(format.date("23.11.2026", "d MMMM yyyy"))
                    .isEqualTo("23 listopada 2026");
        }

        @Test
        @DisplayName("java.util.Date is formatted without shifting the day")
        void utilDate() {
            Date date = Date.from(LocalDate.of(2026, 11, 23)
                    .atStartOfDay(ZoneId.systemDefault())
                    .toInstant());
            assertThat(format.date(date, "dd.MM.yyyy"))
                    .isEqualTo("23.11.2026");
        }

        @Test
        @DisplayName("default pattern when none is given")
        void defaultPattern() {
            assertThat(format.date(LocalDate.of(2026, 11, 23)))
                    .isEqualTo("23.11.2026");
            assertThat(format.date(LocalDate.of(2026, 11, 23), "  "))
                    .isEqualTo("23.11.2026");
        }

        @Test
        @DisplayName("null stays null")
        void nullValue() {
            assertThat(format.date(null, "dd.MM.yyyy")).isNull();
        }

        @Test
        @DisplayName("unparseable value is returned unchanged so the problem stays visible")
        void unparseable() {
            assertThat(format.date("brak daty", "dd.MM.yyyy"))
                    .isEqualTo("brak daty");
        }

        @Test
        @DisplayName("invalid pattern leaves the value unchanged")
        void invalidPattern() {
            assertThat(format.date(LocalDate.of(2026, 11, 23), "dd.MM.yyyy #"))
                    .isEqualTo("2026-11-23");
        }
    }

    // ==================== NUMBER ====================

    @Nested
    @DisplayName("number()")
    class NumberTests {

        @Test
        @DisplayName("pattern with grouping and two decimals, pl_PL separators")
        void grouping() {
            assertThat(normalizeSpaces(format.number(1234.5, "#,##0.00")))
                    .isEqualTo("1 234,50");
        }

        @Test
        @DisplayName("BigDecimal input is formatted exactly")
        void bigDecimal() {
            assertThat(normalizeSpaces(format.number(new BigDecimal("1234567.891"), "#,##0.00")))
                    .isEqualTo("1 234 567,89");
        }

        @Test
        @DisplayName("numeric string with a decimal comma is accepted")
        void numericString() {
            assertThat(normalizeSpaces(format.number("1234,50", "#,##0.00")))
                    .isEqualTo("1 234,50");
        }

        @Test
        @DisplayName("no binary floating point artefacts")
        void noFloatingPointArtifacts() {
            assertThat(format.number(0.1 + 0.2, "0.000"))
                    .isEqualTo("0,300");
        }

        @Test
        @DisplayName("default pattern when none is given")
        void defaultPattern() {
            assertThat(normalizeSpaces(format.number(new BigDecimal("99.9"))))
                    .isEqualTo("99,90");
        }

        @Test
        @DisplayName("null stays null")
        void nullValue() {
            assertThat(format.number(null, "#,##0.00")).isNull();
        }

        @Test
        @DisplayName("non-numeric value is returned unchanged")
        void nonNumeric() {
            assertThat(format.number("nie liczba", "#,##0.00"))
                    .isEqualTo("nie liczba");
        }

        @Test
        @DisplayName("invalid pattern leaves the value unchanged")
        void invalidPattern() {
            assertThat(format.number(new BigDecimal("12.5"), "0.00.00"))
                    .isEqualTo("12.5");
        }
    }

    // ==================== HELPERS ====================

    /** pl_PL uses a no-break space as the grouping separator; normalise it for readable assertions. */
    private static String normalizeSpaces(String value) {
        return value == null ? null : value.replace('\u00A0', ' ').replace('\u202F', ' ');
    }
}
