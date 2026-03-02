package me.muban.docx;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.Locale;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link LocaleUtils}.
 */
@DisplayName("LocaleUtils")
class LocaleUtilsTest {

    @Nested
    @DisplayName("parseLocale")
    class ParseLocale {

        @Test
        @DisplayName("should parse language-only locale")
        void shouldParseLanguageOnly() {
            Locale locale = LocaleUtils.parseLocale("pl");
            assertThat(locale.getLanguage()).isEqualTo("pl");
            assertThat(locale.getCountry()).isEmpty();
        }

        @Test
        @DisplayName("should parse language + country locale")
        void shouldParseLanguageAndCountry() {
            Locale locale = LocaleUtils.parseLocale("pl_PL");
            assertThat(locale.getLanguage()).isEqualTo("pl");
            assertThat(locale.getCountry()).isEqualTo("PL");
        }

        @Test
        @DisplayName("should parse language + country + variant locale")
        void shouldParseFullLocale() {
            Locale locale = LocaleUtils.parseLocale("no_NO_NY");
            assertThat(locale.getLanguage()).isEqualTo("no");
            assertThat(locale.getCountry()).isEqualTo("NO");
            assertThat(locale.getVariant()).isEqualTo("NY");
        }

        @Test
        @DisplayName("should trim whitespace from locale string")
        void shouldTrimWhitespace() {
            Locale locale = LocaleUtils.parseLocale("  en_US  ");
            assertThat(locale.getLanguage()).isEqualTo("en");
            assertThat(locale.getCountry()).isEqualTo("US");
        }

        @Test
        @DisplayName("should throw on null locale string")
        void shouldThrowOnNull() {
            assertThatThrownBy(() -> LocaleUtils.parseLocale(null))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("null or empty");
        }

        @Test
        @DisplayName("should throw on empty locale string")
        void shouldThrowOnEmpty() {
            assertThatThrownBy(() -> LocaleUtils.parseLocale(""))
                    .isInstanceOf(IllegalArgumentException.class);
        }

        @Test
        @DisplayName("should throw on malformed locale with too many parts")
        void shouldThrowOnMalformed() {
            assertThatThrownBy(() -> LocaleUtils.parseLocale("a_b_c_d"))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("Invalid locale format");
        }
    }

    @Nested
    @DisplayName("formatValue")
    class FormatValue {

        @Test
        @DisplayName("should return null for null value")
        void shouldReturnNullForNull() {
            assertThat(LocaleUtils.formatValue(null, Locale.US)).isNull();
            assertThat(LocaleUtils.formatValue(null, null)).isNull();
        }

        @Test
        @DisplayName("should format double with US locale")
        void shouldFormatDoubleUS() {
            assertThat(LocaleUtils.formatValue(1234.56, Locale.US)).isEqualTo("1,234.56");
        }

        @Test
        @DisplayName("should format double with German locale")
        void shouldFormatDoubleGerman() {
            assertThat(LocaleUtils.formatValue(1234.56, Locale.GERMANY)).isEqualTo("1.234,56");
        }

        @Test
        @DisplayName("should format double with Polish locale")
        void shouldFormatDoublePolish() {
            Locale plPL = new Locale("pl", "PL");
            String result = LocaleUtils.formatValue(1234.56, plPL);
            assertThat(result).contains(",");
            assertThat(result).contains("234");
            assertThat(result).contains("56");
        }

        @Test
        @DisplayName("should format integer with locale")
        void shouldFormatInteger() {
            assertThat(LocaleUtils.formatValue(1500, Locale.US)).isEqualTo("1,500");
            assertThat(LocaleUtils.formatValue(42, Locale.US)).isEqualTo("42");
        }

        @Test
        @DisplayName("should format long with locale")
        void shouldFormatLong() {
            assertThat(LocaleUtils.formatValue(1000000L, Locale.US)).isEqualTo("1,000,000");
        }

        @Test
        @DisplayName("should use plain String.valueOf when locale is null")
        void shouldUsePlainWithoutLocale() {
            assertThat(LocaleUtils.formatValue(1234.56, null)).isEqualTo("1234.56");
            assertThat(LocaleUtils.formatValue(42, null)).isEqualTo("42");
        }

        @Test
        @DisplayName("should not format strings even with locale")
        void shouldNotFormatStrings() {
            assertThat(LocaleUtils.formatValue("hello", Locale.US)).isEqualTo("hello");
            assertThat(LocaleUtils.formatValue("ABC-123", Locale.GERMANY)).isEqualTo("ABC-123");
        }

        @Test
        @DisplayName("should not format booleans even with locale")
        void shouldNotFormatBooleans() {
            assertThat(LocaleUtils.formatValue(true, Locale.US)).isEqualTo("true");
            assertThat(LocaleUtils.formatValue(false, Locale.GERMANY)).isEqualTo("false");
        }

        @Test
        @DisplayName("should handle float values")
        void shouldHandleFloat() {
            assertThat(LocaleUtils.formatValue(99.99f, Locale.US)).isEqualTo("99.99");
        }
    }
}
