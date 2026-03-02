package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.sharedtypes.STOnOff;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;

/**
 * Resolves Word table style properties into direct formatting for XSL-FO export.
 *
 * <p>Word table styles can define rich formatting — borders, cell shading, alternating
 * row colors (banding), and text formatting for specific table regions (header row,
 * first column, etc.). The docx4j XSL-FO export pipeline does not resolve style-based
 * table properties, so this utility flattens them into direct properties before PDF export.
 *
 * <h3>Resolved properties:</h3>
 * <ul>
 *   <li><b>Table borders</b> — from {@code Style.getTblPr().getTblBorders()}</li>
 *   <li><b>Cell shading</b> — background colors from default, banding, and positional styles</li>
 *   <li><b>Cell borders</b> — from conditional style overrides</li>
 *   <li><b>Text formatting</b> — bold, italic, color, size from conditional RPr</li>
 * </ul>
 *
 * <h3>Supported conditional types:</h3>
 * <p>firstRow, lastRow, firstCol, lastCol, banding (horizontal and vertical),
 * corner cells (NE, NW, SE, SW), and wholeTable defaults.
 *
 * <h3>Precedence (lowest → highest):</h3>
 * <p>wholeTable → banding → column → row → corner cell → direct formatting.
 * Direct formatting on the cell/run always takes precedence.
 */
public class DocxTableStyleResolver {

    private static final Logger log = LoggerFactory.getLogger(DocxTableStyleResolver.class);

    private DocxTableStyleResolver() {}

    /**
     * Resolve table conditional formatting across the entire document.
     *
     * <p>Walks the document body, finds all tables with style-based conditional formatting,
     * and applies the conditional RPr directly to runs in matching cells.
     *
     * @param wordPackage the DOCX document to process
     */
    public static void resolveTableConditionalFormatting(WordprocessingMLPackage wordPackage) {
        MainDocumentPart mainPart = wordPackage.getMainDocumentPart();
        StyleDefinitionsPart stylesPart;
        try {
            stylesPart = mainPart.getStyleDefinitionsPart();
        } catch (Exception e) {
            log.debug("Could not access style definitions, skipping table style resolution: {}", e.getMessage());
            return;
        }
        if (stylesPart == null) return;

        resolveTableStylesRecursive(mainPart.getContent(), stylesPart);
    }

    private static void resolveTableStylesRecursive(List<Object> content, StyleDefinitionsPart stylesPart) {
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (unwrapped instanceof Tbl table) {
                resolveTableStyles(table, stylesPart);
            } else if (unwrapped instanceof ContentAccessor accessor) {
                resolveTableStylesRecursive(accessor.getContent(), stylesPart);
            }
        }
    }

    /**
     * Resolve all style-based formatting for a single table.
     *
     * <p>Reads the table's style definition and flattens borders, shading, and text
     * formatting into direct properties. Handles table-level borders, default cell
     * properties, banding (alternating row/column colors), and positional conditionals
     * (first row, last column, corner cells, etc.).
     */
    private static void resolveTableStyles(Tbl table, StyleDefinitionsPart stylesPart) {
        TblPr tblPr = table.getTblPr();
        if (tblPr == null || tblPr.getTblStyle() == null) return;

        String styleId = tblPr.getTblStyle().getVal();
        Style tableStyle = stylesPart.getStyleById(styleId);
        if (tableStyle == null) return;

        // 1. Resolve table-level borders from style
        resolveTableBorders(tblPr, tableStyle);

        // 2. Collect conditional overrides (RPr + TcPr) from style
        Map<STTblStyleOverrideType, RPr> conditionalRPr = new EnumMap<>(STTblStyleOverrideType.class);
        Map<STTblStyleOverrideType, TcPr> conditionalTcPr = new EnumMap<>(STTblStyleOverrideType.class);

        List<CTTblStylePr> styleOverrides = tableStyle.getTblStylePr();
        if (styleOverrides != null) {
            for (CTTblStylePr tblStylePr : styleOverrides) {
                if (tblStylePr.getType() == null) continue;
                if (tblStylePr.getRPr() != null) {
                    conditionalRPr.put(tblStylePr.getType(), tblStylePr.getRPr());
                }
                if (tblStylePr.getTcPr() != null) {
                    conditionalTcPr.put(tblStylePr.getType(), tblStylePr.getTcPr());
                }
            }
        }

        // Style's default TcPr applies as baseline for all cells
        TcPr defaultTcPr = tableStyle.getTcPr();

        if (conditionalRPr.isEmpty() && conditionalTcPr.isEmpty() && defaultTcPr == null) return;

        // 3. Determine which conditional types are enabled via tblLook
        CTTblLook tblLook = tblPr.getTblLook();
        boolean firstRowEnabled = isSTOnOffTrue(tblLook != null ? tblLook.getFirstRow() : null);
        boolean lastRowEnabled = isSTOnOffTrue(tblLook != null ? tblLook.getLastRow() : null);
        boolean firstColEnabled = isSTOnOffTrue(tblLook != null ? tblLook.getFirstColumn() : null);
        boolean lastColEnabled = isSTOnOffTrue(tblLook != null ? tblLook.getLastColumn() : null);
        boolean hBandEnabled = !isSTOnOffTrue(tblLook != null ? tblLook.getNoHBand() : null);
        boolean vBandEnabled = !isSTOnOffTrue(tblLook != null ? tblLook.getNoVBand() : null);

        // 4. Collect rows
        List<Tr> rows = new ArrayList<>();
        for (Object rowObj : table.getContent()) {
            Object unwrapped = DocxXmlUtils.unwrap(rowObj);
            if (unwrapped instanceof Tr tr) {
                rows.add(tr);
            }
        }

        if (rows.isEmpty()) return;

        // 5. Apply formatting to each cell based on position
        for (int rowIdx = 0; rowIdx < rows.size(); rowIdx++) {
            boolean isFirstRow = (rowIdx == 0) && firstRowEnabled;
            boolean isLastRow = (rowIdx == rows.size() - 1) && lastRowEnabled;

            boolean isBand1H = false, isBand2H = false;
            if (hBandEnabled && !isFirstRow && !isLastRow) {
                int bandIdx = firstRowEnabled ? rowIdx - 1 : rowIdx;
                if (bandIdx >= 0) {
                    isBand1H = (bandIdx % 2 == 0);
                    isBand2H = (bandIdx % 2 == 1);
                }
            }

            List<Tc> cells = new ArrayList<>();
            for (Object cellObj : rows.get(rowIdx).getContent()) {
                Object unwrapped = DocxXmlUtils.unwrap(cellObj);
                if (unwrapped instanceof Tc tc) {
                    cells.add(tc);
                }
            }

            for (int colIdx = 0; colIdx < cells.size(); colIdx++) {
                boolean isFirstCol = (colIdx == 0) && firstColEnabled;
                boolean isLastCol = (colIdx == cells.size() - 1) && lastColEnabled;

                boolean isBand1V = false, isBand2V = false;
                if (vBandEnabled && !isFirstCol && !isLastCol) {
                    int bandCol = firstColEnabled ? colIdx - 1 : colIdx;
                    if (bandCol >= 0) {
                        isBand1V = (bandCol % 2 == 0);
                        isBand2V = (bandCol % 2 == 1);
                    }
                }

                List<RPr> applicableRPr = collectApplicableRPr(conditionalRPr,
                        isFirstRow, isLastRow, isFirstCol, isLastCol,
                        isBand1H, isBand2H, isBand1V, isBand2V);

                List<TcPr> applicableTcPr = collectApplicableTcPr(conditionalTcPr, defaultTcPr,
                        isFirstRow, isLastRow, isFirstCol, isLastCol,
                        isBand1H, isBand2H, isBand1V, isBand2V);

                Tc cell = cells.get(colIdx);
                if (!applicableRPr.isEmpty()) {
                    applyConditionalFormattingToCell(cell, applicableRPr);
                }
                if (!applicableTcPr.isEmpty()) {
                    applyConditionalCellProperties(cell, applicableTcPr);
                }
            }
        }

        log.debug("Resolved table style '{}' — borders, shading, and text formatting applied", styleId);
    }

    private static void resolveTableBorders(TblPr tblPr, Style tableStyle) {
        CTTblPrBase styleTblPr = tableStyle.getTblPr();
        if (styleTblPr == null || styleTblPr.getTblBorders() == null) return;

        TblBorders styleBorders = styleTblPr.getTblBorders();

        if (tblPr.getTblBorders() == null) {
            tblPr.setTblBorders(styleBorders);
        } else {
            TblBorders direct = tblPr.getTblBorders();
            if (direct.getTop() == null) direct.setTop(styleBorders.getTop());
            if (direct.getBottom() == null) direct.setBottom(styleBorders.getBottom());
            if (direct.getLeft() == null) direct.setLeft(styleBorders.getLeft());
            if (direct.getRight() == null) direct.setRight(styleBorders.getRight());
            if (direct.getInsideH() == null) direct.setInsideH(styleBorders.getInsideH());
            if (direct.getInsideV() == null) direct.setInsideV(styleBorders.getInsideV());
        }
    }

    private static List<RPr> collectApplicableRPr(
            Map<STTblStyleOverrideType, RPr> conditionalRPr,
            boolean isFirstRow, boolean isLastRow, boolean isFirstCol, boolean isLastCol,
            boolean isBand1H, boolean isBand2H, boolean isBand1V, boolean isBand2V) {

        List<RPr> result = new ArrayList<>();

        addIfPresent(result, conditionalRPr, STTblStyleOverrideType.WHOLE_TABLE);

        if (isBand1H) addIfPresent(result, conditionalRPr, STTblStyleOverrideType.BAND_1_HORZ);
        if (isBand2H) addIfPresent(result, conditionalRPr, STTblStyleOverrideType.BAND_2_HORZ);
        if (isBand1V) addIfPresent(result, conditionalRPr, STTblStyleOverrideType.BAND_1_VERT);
        if (isBand2V) addIfPresent(result, conditionalRPr, STTblStyleOverrideType.BAND_2_VERT);

        if (isFirstCol) addIfPresent(result, conditionalRPr, STTblStyleOverrideType.FIRST_COL);
        if (isLastCol)  addIfPresent(result, conditionalRPr, STTblStyleOverrideType.LAST_COL);
        if (isFirstRow) addIfPresent(result, conditionalRPr, STTblStyleOverrideType.FIRST_ROW);
        if (isLastRow)  addIfPresent(result, conditionalRPr, STTblStyleOverrideType.LAST_ROW);

        // Corner cells
        if (isFirstRow && isFirstCol) addIfPresent(result, conditionalRPr, STTblStyleOverrideType.NW_CELL);
        if (isFirstRow && isLastCol)  addIfPresent(result, conditionalRPr, STTblStyleOverrideType.NE_CELL);
        if (isLastRow && isFirstCol)  addIfPresent(result, conditionalRPr, STTblStyleOverrideType.SW_CELL);
        if (isLastRow && isLastCol)   addIfPresent(result, conditionalRPr, STTblStyleOverrideType.SE_CELL);

        return result;
    }

    private static List<TcPr> collectApplicableTcPr(
            Map<STTblStyleOverrideType, TcPr> conditionalTcPr, TcPr defaultTcPr,
            boolean isFirstRow, boolean isLastRow, boolean isFirstCol, boolean isLastCol,
            boolean isBand1H, boolean isBand2H, boolean isBand1V, boolean isBand2V) {

        List<TcPr> result = new ArrayList<>();

        if (defaultTcPr != null) result.add(defaultTcPr);
        addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.WHOLE_TABLE);

        if (isBand1H) addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.BAND_1_HORZ);
        if (isBand2H) addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.BAND_2_HORZ);
        if (isBand1V) addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.BAND_1_VERT);
        if (isBand2V) addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.BAND_2_VERT);

        if (isFirstCol) addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.FIRST_COL);
        if (isLastCol)  addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.LAST_COL);
        if (isFirstRow) addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.FIRST_ROW);
        if (isLastRow)  addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.LAST_ROW);

        // Corner cells
        if (isFirstRow && isFirstCol) addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.NW_CELL);
        if (isFirstRow && isLastCol)  addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.NE_CELL);
        if (isLastRow && isFirstCol)  addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.SW_CELL);
        if (isLastRow && isLastCol)   addIfPresent(result, conditionalTcPr, STTblStyleOverrideType.SE_CELL);

        return result;
    }

    private static <T> void addIfPresent(List<T> list, Map<STTblStyleOverrideType, T> map,
                                          STTblStyleOverrideType type) {
        T value = map.get(type);
        if (value != null) list.add(value);
    }

    private static boolean isSTOnOffTrue(STOnOff value) {
        return value == STOnOff.TRUE || value == STOnOff.ON || value == STOnOff.ONE;
    }

    private static void applyConditionalFormattingToCell(Tc cell, List<RPr> conditionalStyles) {
        for (Object content : cell.getContent()) {
            Object unwrapped = DocxXmlUtils.unwrap(content);
            if (unwrapped instanceof P paragraph) {
                for (Object runObj : paragraph.getContent()) {
                    Object unwrappedRun = DocxXmlUtils.unwrap(runObj);
                    if (unwrappedRun instanceof R run) {
                        RPr runRPr = run.getRPr();
                        if (runRPr == null) {
                            runRPr = new RPr();
                            run.setRPr(runRPr);
                        }
                        mergeConditionalProperties(runRPr, conditionalStyles);
                    }
                }
            }
        }
    }

    private static void mergeConditionalProperties(RPr target, List<RPr> sources) {
        for (RPr source : sources) {
            if (source.getB() != null && target.getB() == null) target.setB(source.getB());
            if (source.getBCs() != null && target.getBCs() == null) target.setBCs(source.getBCs());
            if (source.getI() != null && target.getI() == null) target.setI(source.getI());
            if (source.getICs() != null && target.getICs() == null) target.setICs(source.getICs());
            if (source.getU() != null && target.getU() == null) target.setU(source.getU());
            if (source.getCaps() != null && target.getCaps() == null) target.setCaps(source.getCaps());
            if (source.getSmallCaps() != null && target.getSmallCaps() == null) target.setSmallCaps(source.getSmallCaps());
            if (source.getStrike() != null && target.getStrike() == null) target.setStrike(source.getStrike());
            if (source.getColor() != null && target.getColor() == null) target.setColor(source.getColor());
            if (source.getSz() != null && target.getSz() == null) target.setSz(source.getSz());
            if (source.getSzCs() != null && target.getSzCs() == null) target.setSzCs(source.getSzCs());
        }
    }

    private static void applyConditionalCellProperties(Tc cell, List<TcPr> conditionalStyles) {
        TcPr cellTcPr = cell.getTcPr();
        if (cellTcPr == null) {
            cellTcPr = new TcPr();
            cell.setTcPr(cellTcPr);
        }
        mergeConditionalCellProperties(cellTcPr, conditionalStyles);
    }

    private static void mergeConditionalCellProperties(TcPr target, List<TcPr> sources) {
        CTShd effectiveShd = null;
        TcPrInner.TcBorders effectiveBorders = null;

        for (TcPr source : sources) {
            if (source.getShd() != null) effectiveShd = source.getShd();
            if (source.getTcBorders() != null) effectiveBorders = source.getTcBorders();
        }

        if (effectiveShd != null && target.getShd() == null) target.setShd(effectiveShd);
        if (effectiveBorders != null && target.getTcBorders() == null) target.setTcBorders(effectiveBorders);
    }
}
