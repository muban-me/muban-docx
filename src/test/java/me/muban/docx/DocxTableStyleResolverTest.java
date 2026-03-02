package me.muban.docx;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.sharedtypes.STOnOff;
import org.docx4j.wml.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.math.BigInteger;

import static org.assertj.core.api.Assertions.*;

/**
 * Tests for {@link DocxTableStyleResolver} — table style flattening for XSL-FO export.
 */
@DisplayName("DocxTableStyleResolver")
class DocxTableStyleResolverTest {

    private static final ObjectFactory FACTORY = Context.getWmlObjectFactory();
    private static final String STYLE_ID = "TestTableStyle";

    private WordprocessingMLPackage wordPackage;
    private StyleDefinitionsPart stylesPart;

    @BeforeEach
    void setUp() throws Exception {
        wordPackage = WordprocessingMLPackage.createPackage();
        stylesPart = wordPackage.getMainDocumentPart().getStyleDefinitionsPart();
    }

    // ── Table Borders ──────────────────────────────────────────────────

    @Nested
    @DisplayName("Table Borders")
    class TableBorders {

        @Test
        @DisplayName("should resolve table borders from style to direct TblPr")
        void shouldResolveTableBordersFromStyle() throws Exception {
            TblBorders borders = createTblBorders("4", "single", "000000");
            Style style = createTableStyle(STYLE_ID);
            CTTblPrBase styleTblPr = new CTTblPrBase();
            styleTblPr.setTblBorders(borders);
            style.setTblPr(styleTblPr);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 2, 3, true, false, false, false);
            addTableToDocument(table);

            assertThat(table.getTblPr().getTblBorders()).isNull();

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            TblBorders resolved = table.getTblPr().getTblBorders();
            assertThat(resolved).isNotNull();
            assertThat(resolved.getTop()).isNotNull();
            assertThat(resolved.getTop().getVal()).isEqualTo(STBorder.SINGLE);
            assertThat(resolved.getBottom()).isNotNull();
            assertThat(resolved.getInsideH()).isNotNull();
            assertThat(resolved.getInsideV()).isNotNull();
        }

        @Test
        @DisplayName("should not overwrite direct table borders")
        void shouldNotOverwriteDirectBorders() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            CTTblPrBase styleTblPr = new CTTblPrBase();
            styleTblPr.setTblBorders(createTblBorders("4", "single", "FF0000"));
            style.setTblPr(styleTblPr);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 2, 2, true, false, false, false);

            TblBorders directBorders = new TblBorders();
            CTBorder blueBorder = createBorder("8", STBorder.DOUBLE, "0000FF");
            directBorders.setTop(blueBorder);
            table.getTblPr().setTblBorders(directBorders);

            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            TblBorders resolved = table.getTblPr().getTblBorders();
            assertThat(resolved.getTop().getVal()).isEqualTo(STBorder.DOUBLE);
            assertThat(resolved.getTop().getColor()).isEqualTo("0000FF");
            assertThat(resolved.getBottom()).isNotNull();
            assertThat(resolved.getBottom().getColor()).isEqualTo("FF0000");
        }
    }

    // ── Cell Shading ───────────────────────────────────────────────────

    @Nested
    @DisplayName("Cell Shading")
    class CellShading {

        @Test
        @DisplayName("should apply header row shading from conditional style")
        void shouldApplyHeaderRowShading() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "4472C4", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 2, true, false, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            Tc headerCell = getCell(table, 0, 0);
            assertThat(headerCell.getTcPr()).isNotNull();
            assertThat(headerCell.getTcPr().getShd()).isNotNull();
            assertThat(headerCell.getTcPr().getShd().getFill()).isEqualTo("4472C4");

            Tc dataCell = getCell(table, 1, 0);
            assertThat(dataCell.getTcPr() == null || dataCell.getTcPr().getShd() == null).isTrue();
        }

        @Test
        @DisplayName("should apply last row shading when enabled")
        void shouldApplyLastRowShading() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.LAST_ROW, "E2EFDA", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 2, false, true, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            Tc lastRowCell = getCell(table, 2, 0);
            assertThat(lastRowCell.getTcPr()).isNotNull();
            assertThat(lastRowCell.getTcPr().getShd().getFill()).isEqualTo("E2EFDA");

            Tc middleCell = getCell(table, 1, 0);
            assertThat(middleCell.getTcPr() == null || middleCell.getTcPr().getShd() == null).isTrue();
        }

        @Test
        @DisplayName("should not overwrite direct cell shading")
        void shouldNotOverwriteDirectShading() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "4472C4", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 2, 2, true, false, false, false);

            Tc cell = getCell(table, 0, 0);
            TcPr tcPr = new TcPr();
            CTShd directShd = createShading("FFFF00");
            tcPr.setShd(directShd);
            cell.setTcPr(tcPr);

            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertThat(cell.getTcPr().getShd().getFill()).isEqualTo("FFFF00");
        }

        @Test
        @DisplayName("should apply first column shading")
        void shouldApplyFirstColumnShading() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_COL, "D9E2F3", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 3, false, false, true, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            for (int row = 0; row < 3; row++) {
                Tc firstCol = getCell(table, row, 0);
                assertThat(firstCol.getTcPr()).isNotNull();
                assertThat(firstCol.getTcPr().getShd().getFill()).isEqualTo("D9E2F3");

                Tc secondCol = getCell(table, row, 1);
                assertThat(secondCol.getTcPr() == null || secondCol.getTcPr().getShd() == null).isTrue();
            }
        }
    }

    // ── Banding (Alternating Rows) ─────────────────────────────────────

    @Nested
    @DisplayName("Banding")
    class Banding {

        @Test
        @DisplayName("should apply alternating row shading with header row")
        void shouldApplyAlternatingRowShadingWithHeader() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "4472C4", null);
            addConditionalTcPr(style, STTblStyleOverrideType.BAND_1_HORZ, "D9E2F3", null);
            addConditionalTcPr(style, STTblStyleOverrideType.BAND_2_HORZ, "FFFFFF", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 5, 2, true, false, false, false);
            enableBanding(table, true, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertCellShading(table, 0, 0, "4472C4");
            assertCellShading(table, 1, 0, "D9E2F3");
            assertCellShading(table, 2, 0, "FFFFFF");
            assertCellShading(table, 3, 0, "D9E2F3");
            assertCellShading(table, 4, 0, "FFFFFF");
        }

        @Test
        @DisplayName("should apply alternating row shading without header row")
        void shouldApplyAlternatingRowShadingWithoutHeader() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.BAND_1_HORZ, "D9E2F3", null);
            addConditionalTcPr(style, STTblStyleOverrideType.BAND_2_HORZ, "FFFFFF", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 4, 2, false, false, false, false);
            enableBanding(table, true, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertCellShading(table, 0, 0, "D9E2F3");
            assertCellShading(table, 1, 0, "FFFFFF");
            assertCellShading(table, 2, 0, "D9E2F3");
            assertCellShading(table, 3, 0, "FFFFFF");
        }

        @Test
        @DisplayName("should not apply banding when noHBand is true")
        void shouldNotApplyBandingWhenDisabled() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.BAND_1_HORZ, "D9E2F3", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 2, false, false, false, false);
            CTTblLook look = table.getTblPr().getTblLook();
            look.setNoHBand(STOnOff.TRUE);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            for (int row = 0; row < 3; row++) {
                Tc cell = getCell(table, row, 0);
                assertThat(cell.getTcPr() == null || cell.getTcPr().getShd() == null)
                        .as("Row %d should have no shading", row).isTrue();
            }
        }

        @Test
        @DisplayName("should skip first and last row from banding when they have their own styles")
        void shouldSkipFirstAndLastRowFromBanding() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "HEADER", null);
            addConditionalTcPr(style, STTblStyleOverrideType.LAST_ROW, "FOOTER", null);
            addConditionalTcPr(style, STTblStyleOverrideType.BAND_1_HORZ, "BAND1", null);
            addConditionalTcPr(style, STTblStyleOverrideType.BAND_2_HORZ, "BAND2", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 4, 2, true, true, false, false);
            enableBanding(table, true, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertCellShading(table, 0, 0, "HEADER");
            assertCellShading(table, 1, 0, "BAND1");
            assertCellShading(table, 2, 0, "BAND2");
            assertCellShading(table, 3, 0, "FOOTER");
        }
    }

    // ── Precedence ─────────────────────────────────────────────────────

    @Nested
    @DisplayName("Precedence")
    class PrecedenceTests {

        @Test
        @DisplayName("row conditional should override column conditional for shading")
        void rowShouldOverrideColumnShading() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_COL, "COL_COLOR", null);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "ROW_COLOR", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 3, true, false, true, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertCellShading(table, 0, 0, "ROW_COLOR");
            assertCellShading(table, 1, 0, "COL_COLOR");
        }

        @Test
        @DisplayName("banding should override default but not position conditionals")
        void bandingShouldOverrideDefaultOnly() throws Exception {
            Style style = createTableStyle(STYLE_ID);

            TcPr defaultTcPr = new TcPr();
            defaultTcPr.setShd(createShading("DEFAULT"));
            style.setTcPr(defaultTcPr);

            addConditionalTcPr(style, STTblStyleOverrideType.BAND_1_HORZ, "BAND1", null);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "HEADER", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 2, true, false, false, false);
            enableBanding(table, true, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertCellShading(table, 0, 0, "HEADER");
            assertCellShading(table, 1, 0, "BAND1");
            assertCellShading(table, 2, 0, "DEFAULT");
        }
    }

    // ── Text Formatting (RPr) ──────────────────────────────────────────

    @Nested
    @DisplayName("Text Formatting (RPr)")
    class TextFormatting {

        @Test
        @DisplayName("should apply bold from header row style")
        void shouldApplyBoldFromHeaderStyle() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalRPr(style, STTblStyleOverrideType.FIRST_ROW);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 2, 2, true, false, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            R headerRun = getFirstRun(table, 0, 0);
            assertThat(headerRun.getRPr()).isNotNull();
            assertThat(headerRun.getRPr().getB().isVal()).isTrue();

            R dataRun = getFirstRun(table, 1, 0);
            assertThat(dataRun.getRPr() == null || dataRun.getRPr().getB() == null).isTrue();
        }

        @Test
        @DisplayName("should not overwrite direct text formatting")
        void shouldNotOverwriteDirectFormatting() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalRPr(style, STTblStyleOverrideType.FIRST_ROW);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 2, 2, true, false, false, false);

            R run = getFirstRun(table, 0, 0);
            RPr rpr = new RPr();
            BooleanDefaultTrue italic = new BooleanDefaultTrue();
            italic.setVal(true);
            rpr.setI(italic);
            BooleanDefaultTrue notBold = new BooleanDefaultTrue();
            notBold.setVal(false);
            rpr.setB(notBold);
            run.setRPr(rpr);

            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertThat(run.getRPr().getB().isVal()).isFalse();
            assertThat(run.getRPr().getI().isVal()).isTrue();
        }
    }

    // ── Edge Cases ─────────────────────────────────────────────────────

    @Nested
    @DisplayName("Edge Cases")
    class EdgeCases {

        @Test
        @DisplayName("should handle table with no style")
        void shouldHandleNoStyle() throws Exception {
            Tbl table = createSimpleTable(2, 2);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);
        }

        @Test
        @DisplayName("should handle table with non-existent style reference")
        void shouldHandleNonExistentStyle() throws Exception {
            Tbl table = createStyledTable("NonExistentStyle", 2, 2,
                    true, false, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);
        }

        @Test
        @DisplayName("should handle empty table")
        void shouldHandleEmptyTable() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "4472C4", null);
            addStyleToDocument(style);

            Tbl table = FACTORY.createTbl();
            TblPr tblPr = new TblPr();
            CTTblPrBase.TblStyle tblStyle = new CTTblPrBase.TblStyle();
            tblStyle.setVal(STYLE_ID);
            tblPr.setTblStyle(tblStyle);
            table.setTblPr(tblPr);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);
        }

        @Test
        @DisplayName("should handle single-row table as both first and last")
        void shouldHandleSingleRowTable() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "HEADER", null);
            addConditionalTcPr(style, STTblStyleOverrideType.LAST_ROW, "FOOTER", null);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 1, 2, true, true, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertCellShading(table, 0, 0, "FOOTER");
        }

        @Test
        @DisplayName("should handle style with no conditional overrides but with default TcPr")
        void shouldHandleDefaultTcPrOnly() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            TcPr defaultTcPr = new TcPr();
            defaultTcPr.setShd(createShading("E0E0E0"));
            style.setTcPr(defaultTcPr);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 2, 2, false, false, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            for (int row = 0; row < 2; row++) {
                for (int col = 0; col < 2; col++) {
                    assertCellShading(table, row, col, "E0E0E0");
                }
            }
        }
    }

    // ── Cell Borders ───────────────────────────────────────────────────

    @Nested
    @DisplayName("Cell Borders")
    class CellBorders {

        @Test
        @DisplayName("should apply cell borders from conditional style")
        void shouldApplyCellBordersFromConditionalStyle() throws Exception {
            Style style = createTableStyle(STYLE_ID);
            TcPrInner.TcBorders headerBorders = new TcPrInner.TcBorders();
            headerBorders.setBottom(createBorder("8", STBorder.SINGLE, "000000"));
            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, null, headerBorders);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 2, true, false, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            Tc headerCell = getCell(table, 0, 0);
            assertThat(headerCell.getTcPr()).isNotNull();
            assertThat(headerCell.getTcPr().getTcBorders()).isNotNull();
            assertThat(headerCell.getTcPr().getTcBorders().getBottom().getSz())
                    .isEqualTo(BigInteger.valueOf(8));

            Tc dataCell = getCell(table, 1, 0);
            assertThat(dataCell.getTcPr() == null || dataCell.getTcPr().getTcBorders() == null).isTrue();
        }
    }

    // ── Combined Formatting ────────────────────────────────────────────

    @Nested
    @DisplayName("Combined Formatting")
    class CombinedFormatting {

        @Test
        @DisplayName("should resolve borders, shading, and bold together")
        void shouldResolveBordersShadingAndBold() throws Exception {
            Style style = createTableStyle(STYLE_ID);

            CTTblPrBase styleTblPr = new CTTblPrBase();
            styleTblPr.setTblBorders(createTblBorders("4", "single", "999999"));
            style.setTblPr(styleTblPr);

            addConditionalTcPr(style, STTblStyleOverrideType.FIRST_ROW, "4472C4", null);
            addConditionalRPr(style, STTblStyleOverrideType.FIRST_ROW);
            addStyleToDocument(style);

            Tbl table = createStyledTable(STYLE_ID, 3, 2, true, false, false, false);
            addTableToDocument(table);

            DocxTableStyleResolver.resolveTableConditionalFormatting(wordPackage);

            assertThat(table.getTblPr().getTblBorders()).isNotNull();
            assertThat(table.getTblPr().getTblBorders().getTop().getVal()).isEqualTo(STBorder.SINGLE);

            assertCellShading(table, 0, 0, "4472C4");

            R headerRun = getFirstRun(table, 0, 0);
            assertThat(headerRun.getRPr().getB().isVal()).isTrue();

            Tc dataCell = getCell(table, 1, 0);
            assertThat(dataCell.getTcPr() == null || dataCell.getTcPr().getShd() == null).isTrue();
            R dataRun = getFirstRun(table, 1, 0);
            assertThat(dataRun.getRPr() == null || dataRun.getRPr().getB() == null).isTrue();
        }
    }

    // ── Helper methods ─────────────────────────────────────────────────

    private Style createTableStyle(String styleId) {
        Style style = FACTORY.createStyle();
        style.setStyleId(styleId);
        style.setType("table");
        Style.Name name = new Style.Name();
        name.setVal(styleId);
        style.setName(name);
        return style;
    }

    private void addStyleToDocument(Style style) {
        stylesPart.getJaxbElement().getStyle().add(style);
    }

    private void addTableToDocument(Tbl table) {
        wordPackage.getMainDocumentPart().getContent().add(table);
    }

    private Tbl createStyledTable(String styleId, int rows, int cols,
                                   boolean firstRow, boolean lastRow,
                                   boolean firstCol, boolean lastCol) {
        Tbl table = FACTORY.createTbl();
        TblPr tblPr = new TblPr();

        CTTblPrBase.TblStyle tblStyle = new CTTblPrBase.TblStyle();
        tblStyle.setVal(styleId);
        tblPr.setTblStyle(tblStyle);

        CTTblLook tblLook = new CTTblLook();
        tblLook.setFirstRow(firstRow ? STOnOff.TRUE : STOnOff.FALSE);
        tblLook.setLastRow(lastRow ? STOnOff.TRUE : STOnOff.FALSE);
        tblLook.setFirstColumn(firstCol ? STOnOff.TRUE : STOnOff.FALSE);
        tblLook.setLastColumn(lastCol ? STOnOff.TRUE : STOnOff.FALSE);
        tblLook.setNoHBand(STOnOff.TRUE);
        tblLook.setNoVBand(STOnOff.TRUE);
        tblPr.setTblLook(tblLook);

        table.setTblPr(tblPr);

        for (int r = 0; r < rows; r++) {
            Tr row = FACTORY.createTr();
            for (int c = 0; c < cols; c++) {
                Tc cell = FACTORY.createTc();
                P p = FACTORY.createP();
                R run = FACTORY.createR();
                Text text = FACTORY.createText();
                text.setValue("R" + r + "C" + c);
                run.getContent().add(text);
                p.getContent().add(run);
                cell.getContent().add(p);
                row.getContent().add(cell);
            }
            table.getContent().add(row);
        }

        return table;
    }

    private Tbl createSimpleTable(int rows, int cols) {
        Tbl table = FACTORY.createTbl();
        for (int r = 0; r < rows; r++) {
            Tr row = FACTORY.createTr();
            for (int c = 0; c < cols; c++) {
                Tc cell = FACTORY.createTc();
                P p = FACTORY.createP();
                R run = FACTORY.createR();
                Text text = FACTORY.createText();
                text.setValue("cell");
                run.getContent().add(text);
                p.getContent().add(run);
                cell.getContent().add(p);
                row.getContent().add(cell);
            }
            table.getContent().add(row);
        }
        return table;
    }

    private void enableBanding(Tbl table, boolean horizontal, boolean vertical) {
        CTTblLook look = table.getTblPr().getTblLook();
        look.setNoHBand(horizontal ? STOnOff.FALSE : STOnOff.TRUE);
        look.setNoVBand(vertical ? STOnOff.FALSE : STOnOff.TRUE);
    }

    private void addConditionalTcPr(Style style, STTblStyleOverrideType type,
                                     String shadingFill,
                                     TcPrInner.TcBorders borders) {
        CTTblStylePr override = findOrCreateOverride(style, type);
        TcPr tcPr = override.getTcPr();
        if (tcPr == null) {
            tcPr = new TcPr();
            override.setTcPr(tcPr);
        }
        if (shadingFill != null) {
            tcPr.setShd(createShading(shadingFill));
        }
        if (borders != null) {
            tcPr.setTcBorders(borders);
        }
    }

    private void addConditionalRPr(Style style, STTblStyleOverrideType type) {
        CTTblStylePr override = findOrCreateOverride(style, type);
        RPr rpr = new RPr();
        BooleanDefaultTrue bold = new BooleanDefaultTrue();
        bold.setVal(true);
        rpr.setB(bold);
        override.setRPr(rpr);
    }

    private CTTblStylePr findOrCreateOverride(Style style, STTblStyleOverrideType type) {
        if (style.getTblStylePr() != null) {
            for (CTTblStylePr existing : style.getTblStylePr()) {
                if (existing.getType() == type) return existing;
            }
        }
        CTTblStylePr override = new CTTblStylePr();
        override.setType(type);
        style.getTblStylePr().add(override);
        return override;
    }

    private CTShd createShading(String fill) {
        CTShd shd = new CTShd();
        shd.setVal(STShd.CLEAR);
        shd.setColor("auto");
        shd.setFill(fill);
        return shd;
    }

    private CTBorder createBorder(String size, STBorder style, String color) {
        CTBorder border = new CTBorder();
        border.setVal(style);
        border.setSz(new BigInteger(size));
        border.setColor(color);
        border.setSpace(BigInteger.ZERO);
        return border;
    }

    private TblBorders createTblBorders(String size, String styleStr, String color) {
        STBorder borderStyle = STBorder.SINGLE;
        TblBorders borders = new TblBorders();
        borders.setTop(createBorder(size, borderStyle, color));
        borders.setBottom(createBorder(size, borderStyle, color));
        borders.setLeft(createBorder(size, borderStyle, color));
        borders.setRight(createBorder(size, borderStyle, color));
        borders.setInsideH(createBorder(size, borderStyle, color));
        borders.setInsideV(createBorder(size, borderStyle, color));
        return borders;
    }

    private Tc getCell(Tbl table, int rowIdx, int colIdx) {
        Tr row = (Tr) table.getContent().get(rowIdx);
        return (Tc) row.getContent().get(colIdx);
    }

    private R getFirstRun(Tbl table, int rowIdx, int colIdx) {
        Tc cell = getCell(table, rowIdx, colIdx);
        P p = (P) cell.getContent().get(0);
        return (R) p.getContent().get(0);
    }

    private void assertCellShading(Tbl table, int rowIdx, int colIdx, String expectedFill) {
        Tc cell = getCell(table, rowIdx, colIdx);
        assertThat(cell.getTcPr())
                .as("TcPr for cell (%d,%d)", rowIdx, colIdx).isNotNull();
        assertThat(cell.getTcPr().getShd())
                .as("Shading for cell (%d,%d)", rowIdx, colIdx).isNotNull();
        assertThat(cell.getTcPr().getShd().getFill())
                .as("Shading fill for cell (%d,%d)", rowIdx, colIdx)
                .isEqualTo(expectedFill);
    }
}
