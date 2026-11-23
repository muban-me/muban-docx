package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Tr;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import static me.muban.docx.DocxTestSupport.createDocx;
import static me.muban.docx.DocxTestSupport.createParagraph;
import static me.muban.docx.DocxTestSupport.extractTexts;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * End-to-end tests for template variables ({@code #{let}}) and the {@code #fmt} helper running
 * through the real processing pipeline.
 *
 * <p>These cover the combination that the individual unit tests cannot: declarations resolved
 * before conditional blocks, `${name}` substitution afterwards, and {@code #fmt} reaching both a
 * branch condition and a replicated table row.
 */
@DisplayName("MubanDocxEngine — variables and formatting")
class MubanDocxEngineVariablesTest {

    private static final Locale PL = Locale.forLanguageTag("pl-PL");

    private static final String SALUTATION_DECLARATION =
            "#{let salutation = addressee_gender == 'M' ? 'Szanowny Panie,' "
            + ": (addressee_gender == 'F' ? 'Szanowna Pani,' : 'Szanowni Państwo,')}";

    private static List<String> processAndRead(WordprocessingMLPackage docx, Map<String, Object> context) {
        WordprocessingMLPackage processed = MubanDocxEngine.builder()
                .template(docx)
                .rawContext(context)
                .locale(PL)
                .build()
                .process();
        return extractTexts(processed.getMainDocumentPart().getContent()).stream()
                .map(MubanDocxEngineVariablesTest::normalizeSpaces)
                .toList();
    }

    // ==================== LETTER SCENARIO ====================

    @Nested
    @DisplayName("letter assembled from declarations")
    class LetterScenario {

        @Test
        @DisplayName("gender, date and amount are formatted by the template patterns")
        void fullPipeline() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    SALUTATION_DECLARATION,
                    "#{let due = #fmt.date(document_date, 'd MMMM yyyy')}",
                    "${salutation}",
                    "Termin płatności: ${due}.",
                    "${#fmt.number(debt_total_amount, '#,##0.00')} zł");

            Map<String, Object> context = new LinkedHashMap<>();
            context.put("addressee_gender", "F");
            context.put("document_date", "2026-11-23");
            context.put("debt_total_amount", new BigDecimal("1234.5"));

            List<String> texts = processAndRead(docx, context);

            assertThat(texts).containsExactly(
                    "Szanowna Pani,",
                    "Termin płatności: 23 listopada 2026.",
                    "1 234,50 zł");
        }

        @Test
        @DisplayName("the masculine variant is selected from the same declaration")
        void masculineVariant() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    SALUTATION_DECLARATION,
                    "${salutation}");

            List<String> texts = processAndRead(docx, Map.of("addressee_gender", "M"));

            assertThat(texts).containsExactly("Szanowny Panie,");
        }

        @Test
        @DisplayName("a company falls back to the neutral variant")
        void companyVariant() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    SALUTATION_DECLARATION,
                    "${salutation}");

            List<String> texts = processAndRead(docx, Map.of("addressee_gender", "CORPORATE"));

            assertThat(texts).containsExactly("Szanowni Państwo,");
        }
    }

    // ==================== CONDITIONS ====================

    @Nested
    @DisplayName("variables in conditional blocks")
    class ConditionalIntegration {

        @Test
        @DisplayName("a branch condition may reference a declared variable")
        void variableInCondition() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{let is_company = addressee_gender != 'M' and addressee_gender != 'F'}",
                    "#{if is_company}",
                    "Państwa zobowiązanie",
                    "#{else}",
                    "Pana/Pani zobowiązanie",
                    "#{fi}");

            List<String> texts = processAndRead(docx, Map.of("addressee_gender", "CORPORATE"));

            assertThat(texts).containsExactly("Państwa zobowiązanie");
        }

        @Test
        @DisplayName("#fmt is reachable from a condition")
        void formatHelperInCondition() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    "#{if #fmt.date(document_date, 'yyyy') == '2026'}",
                    "Rok bieżący",
                    "#{else}",
                    "Inny rok",
                    "#{fi}");

            List<String> texts = processAndRead(docx, Map.of("document_date", "2026-11-23"));

            assertThat(texts).containsExactly("Rok bieżący");
        }
    }

    // ==================== TABLE ROWS ====================

    @Nested
    @DisplayName("table row replication")
    class TableIntegration {

        @Test
        @DisplayName("a row wrapped in an expression is still recognised as array-bound and formatted")
        void fmtInsideReplicatedRow() throws Exception {
            Tbl table = tableWithTemplateRow(
                    "${#fmt.date(installments.due_date, 'dd.MM.yyyy')}");

            WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
            docx.getMainDocumentPart().getContent().clear();
            docx.getMainDocumentPart().getContent().add(table);

            Map<String, Object> context = new LinkedHashMap<>();
            context.put("installments", List.of(
                    Map.of("due_date", "2026-12-01"),
                    Map.of("due_date", "2027-01-01")));

            WordprocessingMLPackage processed = MubanDocxEngine.builder()
                    .template(docx)
                    .rawContext(context)
                    .locale(PL)
                    .build()
                    .process();

            Tbl processedTable = (Tbl) DocxXmlUtils.unwrap(processed.getMainDocumentPart().getContent().get(0));

            assertThat(rowTexts(processedTable))
                    .containsExactly("Termin", "01.12.2026", "01.01.2027");
        }
    }

    // ==================== DISCOVERY AND FAILURE ====================

    @Nested
    @DisplayName("template introspection and failure visibility")
    class Introspection {

        @Test
        @DisplayName("declarations can be listed before processing")
        void extractDeclarations() throws Exception {
            WordprocessingMLPackage docx = createDocx(
                    SALUTATION_DECLARATION,
                    "#{let due = #fmt.date(document_date, 'd MMMM yyyy')}",
                    "${salutation}");

            Map<String, String> declared = MubanDocxEngine.extractDeclaredVariables(docx);

            assertThat(declared).containsOnlyKeys("salutation", "due");
            assertThat(declared.get("due")).isEqualTo("#fmt.date(document_date, 'd MMMM yyyy')");
        }

        @Test
        @DisplayName("an undefined variable stays visible in the output")
        void undefinedVariableStaysVisible() throws Exception {
            WordprocessingMLPackage docx = createDocx("Wartość: ${nie_ma_takiego_pola}");

            List<String> texts = processAndRead(docx, Map.of());

            assertThat(texts).containsExactly("Wartość: ${nie_ma_takiego_pola}");
        }
    }

    // ==================== HELPERS ====================

    private static Tbl tableWithTemplateRow(String templateCellText) {
        Tr header = new Tr();
        Tc headerCell = new Tc();
        headerCell.getContent().add(createParagraph("Termin"));
        header.getContent().add(headerCell);

        Tr template = new Tr();
        Tc templateCell = new Tc();
        templateCell.getContent().add(createParagraph(templateCellText));
        template.getContent().add(templateCell);

        Tbl table = new Tbl();
        table.getContent().add(header);
        table.getContent().add(template);
        return table;
    }

    private static List<String> rowTexts(Tbl table) {
        List<String> rows = new ArrayList<>();
        for (Object rowObj : table.getContent()) {
            Object row = DocxXmlUtils.unwrap(rowObj);
            if (!(row instanceof Tr tr)) {
                continue;
            }
            StringBuilder text = new StringBuilder();
            for (Object cellObj : tr.getContent()) {
                Object cell = DocxXmlUtils.unwrap(cellObj);
                if (cell instanceof Tc tc) {
                    for (Object paragraphObj : tc.getContent()) {
                        Object paragraph = DocxXmlUtils.unwrap(paragraphObj);
                        if (paragraph instanceof P) {
                            text.append(DocxXmlUtils.extractAllText(List.of(paragraph)));
                        }
                    }
                }
            }
            rows.add(text.toString());
        }
        return rows;
    }

    private static String normalizeSpaces(String value) {
        return value.replace('\u00A0', ' ').replace('\u202F', ' ');
    }
}
