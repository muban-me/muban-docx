package me.muban.docx;

import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import static me.muban.docx.DocxTestSupport.createParagraph;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for {@link DocxVariableProcessor} — template variables declared with
 * {@code #{let name = expression}}.
 */
@DisplayName("DocxVariableProcessor")
class DocxVariableProcessorTest {

    private static final Locale PL = Locale.forLanguageTag("pl-PL");

    private Map<String, Object> context;
    private Map<String, Object> variables;

    @BeforeEach
    void setUp() {
        context = new LinkedHashMap<>();
        variables = new LinkedHashMap<>();
        variables.put(MubanDocxEngine.FORMAT_HELPER_VARIABLE, new DocxFormatHelper(PL));
    }

    private int process(List<Object> content) {
        return DocxVariableProcessor.processVariables(content, context, variables, PL);
    }

    private static List<Object> contentOf(String... paragraphTexts) {
        List<Object> content = new ArrayList<>();
        for (String text : paragraphTexts) {
            content.add(createParagraph(text));
        }
        return content;
    }

    // ==================== DECLARATIONS ====================

    @Nested
    @DisplayName("declaration handling")
    class DeclarationTests {

        @Test
        @DisplayName("declares a value and removes the marker paragraph")
        void declaresAndRemovesMarker() {
            List<Object> content = contentOf("#{let salutation = 'Szanowny Panie,'}");

            int resolved = process(content);

            assertThat(resolved).isEqualTo(1);
            assertThat(variables).containsEntry("salutation", "Szanowny Panie,");
            assertThat(content).isEmpty();
        }

        @Test
        @DisplayName("keeps surrounding paragraphs intact")
        void keepsSurroundingContent() {
            List<Object> content = contentOf(
                    "Przed",
                    "#{let amount = 1500}",
                    "Po");

            process(content);

            assertThat(DocxTestSupport.extractTexts(content))
                    .containsExactly("Przed", "Po");
        }

        @Test
        @DisplayName("a later declaration sees the earlier ones")
        void sequentialEvaluation() {
            List<Object> content = contentOf(
                    "#{let base = 100}",
                    "#{let with_interest = #base * 2}");

            process(content);

            assertThat(variables).containsEntry("base", 100);
            assertThat(variables).containsEntry("with_interest", 200);
        }

        @Test
        @DisplayName("a forward reference leaves the variable undefined")
        void forwardReference() {
            List<Object> content = contentOf("#{let doubled = #base * 2}");

            int resolved = process(content);

            assertThat(resolved).isZero();
            assertThat(variables).doesNotContainKey("doubled");
            assertThat(content).isEmpty();
        }

        @Test
        @DisplayName("a failing expression is removed but leaves the variable undefined")
        void failedExpression() {
            List<Object> content = contentOf("#{let broken = no_such_field}");

            assertThat(process(content)).isZero();
            assertThat(variables).doesNotContainKey("broken");
            assertThat(content).isEmpty();
        }

        @Test
        @DisplayName("a declaration never writes into the input data namespace")
        void doesNotTouchInputData() {
            context.put("addressee_name", "z UDA");
            List<Object> content = contentOf("#{let addressee_name = 'z szablonu'}");

            process(content);

            assertThat(variables).containsEntry("addressee_name", "z szablonu");
            assertThat(context).containsEntry("addressee_name", "z UDA");
        }

        @Test
        @DisplayName("a declaration cannot take a name reserved by SpEL")
        void reservedName() {
            List<Object> content = contentOf("#{let root = 'próba'}");

            assertThat(process(content)).isZero();
            assertThat(variables).doesNotContainKey("root");
            assertThat(content).isEmpty();
        }

        @Test
        @DisplayName("a declaration cannot replace an engine helper")
        void helperNameReserved() {
            List<Object> content = contentOf("#{let fmt = 'próba'}");

            assertThat(process(content)).isZero();
            assertThat(variables.get("fmt")).isInstanceOf(DocxFormatHelper.class);
            assertThat(content).isEmpty();
        }

        @Test
        @DisplayName("the raw value type is preserved, not stringified")
        void keepsRawType() {
            List<Object> content = contentOf("#{let amount = debt * 2}");
            context.put("debt", new java.math.BigDecimal("12.50"));

            process(content);

            assertThat(variables.get("amount")).isInstanceOf(java.math.BigDecimal.class);
            assertThat(variables.get("amount")).isEqualTo(new java.math.BigDecimal("25.00"));
        }

        @Test
        @DisplayName("null is a legitimate value and defines the variable")
        void nullValueDefinesVariable() {
            context.put("nothing", null);
            List<Object> content = contentOf("#{let copy = nothing}");

            process(content);

            assertThat(variables).containsKey("copy");
            assertThat(variables.get("copy")).isNull();
        }

        @Test
        @DisplayName("a malformed marker is logged and left in place")
        void malformedMarker() {
            List<Object> content = contentOf("#{let brak_rownania}");

            assertThat(process(content)).isZero();
            assertThat(content).hasSize(1);
        }

        @Test
        @DisplayName("an expression may contain a closing brace")
        void expressionWithBrace() {
            List<Object> content = contentOf("#{let literal = '#{}'}");

            process(content);

            assertThat(variables).containsEntry("literal", "#{}");
        }

        @Test
        @DisplayName("markers split across Word runs are rejoined")
        void splitRuns() {
            P paragraph = new P();
            paragraph.getContent().add(run("#{let x = 'A"));
            paragraph.getContent().add(run("B'}"));
            List<Object> content = new ArrayList<>();
            content.add(paragraph);

            process(content);

            assertThat(variables).containsEntry("x", "AB");
            assertThat(content).isEmpty();
        }

        @Test
        @DisplayName("declarations nested in a table are resolved too")
        void nestedInTable() {
            Tc cell = new Tc();
            cell.getContent().add(createParagraph("#{let inner = 'z tabeli'}"));
            Tr row = new Tr();
            row.getContent().add(cell);
            Tbl table = new Tbl();
            table.getContent().add(row);

            List<Object> content = new ArrayList<>();
            content.add(table);

            assertThat(process(content)).isEqualTo(1);
            assertThat(variables).containsEntry("inner", "z tabeli");
            assertThat(cell.getContent()).isEmpty();
        }

        @Test
        @DisplayName("gender declension is an expression in the declaration, not a library feature")
        void genderDeclension() {
            context.put("addressee_gender", "F");
            List<Object> content = contentOf(
                    "#{let salutation = addressee_gender == 'M' ? 'Szanowny Panie,' : (addressee_gender == 'F' ? 'Szanowna Pani,' : 'Szanowni Państwo,')}");

            process(content);

            assertThat(variables).containsEntry("salutation", "Szanowna Pani,");
        }
    }

    // ==================== HELPERS IN DECLARATIONS ====================

    @Nested
    @DisplayName("format helper inside a declaration")
    class HelperTests {

        @Test
        @DisplayName("#fmt is reachable from a declaration")
        void formatHelperAvailable() {
            context.put("document_date", "2026-11-23");
            List<Object> content = contentOf("#{let due = #fmt.date(document_date, 'd MMMM yyyy')}");

            process(content);

            assertThat(variables).containsEntry("due", "23 listopada 2026");
        }
    }

    // ==================== EXTRACTION ====================

    @Nested
    @DisplayName("collectDeclarations()")
    class ExtractionTests {

        @Test
        @DisplayName("returns names and expressions in document order")
        void inDocumentOrder() {
            Map<String, String> declarations = DocxVariableProcessor.newDeclarationMap();
            List<Object> content = contentOf(
                    "#{let first = 1}",
                    "Text",
                    "#{let second = first + 1}");

            DocxVariableProcessor.collectDeclarations(content, declarations);

            assertThat(declarations).containsExactly(
                    Map.entry("first", "1"),
                    Map.entry("second", "first + 1"));
        }

        @Test
        @DisplayName("scans nested content")
        void nestedContent() {
            Tc cell = new Tc();
            cell.getContent().add(createParagraph("#{let nested = 'x'}"));
            Tr row = new Tr();
            row.getContent().add(cell);
            Tbl table = new Tbl();
            table.getContent().add(row);

            Map<String, String> declarations = DocxVariableProcessor.newDeclarationMap();
            DocxVariableProcessor.collectDeclarations(List.of(table), declarations);

            assertThat(declarations).containsOnlyKeys("nested");
        }

        @Test
        @DisplayName("empty when the template declares nothing")
        void nothingDeclared() {
            Map<String, String> declarations = DocxVariableProcessor.newDeclarationMap();

            DocxVariableProcessor.collectDeclarations(contentOf("Text", "${value}"), declarations);

            assertThat(declarations).isEmpty();
        }
    }

    private static R run(String text) {
        R run = new R();
        Text value = new Text();
        value.setValue(text);
        value.setSpace("preserve");
        run.getContent().add(value);
        return run;
    }
}
