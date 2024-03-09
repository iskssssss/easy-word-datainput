package top.kongsheng.common.word.datainput.utils;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.style.CellStyle;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.model.ParagraphCreateFun;
import top.kongsheng.common.word.datainput.config.FullConfig;
import top.kongsheng.common.word.datainput.model.input.AbsDataInput;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.*;

/**
 * 表格工具类
 *
 * @author 孔胜
 * @date 2023/7/24 16:14
 */
public class TableUtil {

    public static String getTableName(XWPFTable table) {
        CTString tblCaption = table.getCTTbl().getTblPr().getTblCaption();
        if (tblCaption == null) {
            return "";
        }
        return tblCaption.getVal();
    }

    public static String getTableDescription(XWPFTable table) {
        CTString tblDescription = table.getCTTbl().getTblPr().getTblDescription();
        if (tblDescription == null) {
            return "";
        }
        return tblDescription.getVal();
    }

    public static CTTblCellMar getTableCellMar(XWPFTable table) {
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        return tblPr.getTblCellMar();
    }

    public static void clearTableName(XWPFTable table) {
        table.getCTTbl().getTblPr().setTblCaption(EMPTY_STRING);
    }

    public static void clearTableDescription(XWPFTable table) {
        table.getCTTbl().getTblPr().setTblDescription(EMPTY_STRING);
    }

    public static Map<String, XWPFTable> findTable(XWPFDocument document) {
        Map<String, XWPFTable> tableMap = new LinkedHashMap<>();
        List<XWPFTable> tables = handleTable(document.getTables());
        for (XWPFTable table : tables) {
            String tableName = getTableName(table);
            if (StrUtil.isEmpty(tableName)) {
                continue;
            }
            tableMap.put(tableName, table);
        }
        return tableMap;
    }

    public static List<XWPFTable> findTableList(XWPFDocument document) {
        return handleTable(document.getTables());
    }

    public static XWPFTable findTableByName(XWPFDocument document, String tableName) {
        return findTableByName(findTableList(document), tableName);
    }

    public static XWPFTable findTableByName(List<XWPFTable> tableList, String tableName) {
        for (XWPFTable xwpfTable : tableList) {
            if (tableName.equals(TableUtil.getTableName(xwpfTable))) {
                return xwpfTable;
            }
        }
        return null;
    }

    public static List<XWPFTable> handleTable(List<XWPFTable> tables) {
        List<XWPFTable> result = new LinkedList<>();
        for (XWPFTable table : tables) {
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell tableCell : tableCells) {
                    List<XWPFTable> tableList = tableCell.getTables();
                    if (tableList.isEmpty()) {
                        continue;
                    }
                    result.addAll(handleTable(tableList));
                }
            }
            result.add(table);
        }
        return result;
    }

    public static XWPFTableRow findRow(XWPFTable table, int rowIndex) {
        XWPFTableRow tableRow = table.getRow(rowIndex);
        XWPFTableRow row = tableRow == null ? table.createRow() : tableRow;
        return row;
    }

    public static boolean isEmpty(XWPFTable table, int rowIndex) {
        XWPFTableRow tableRow = table.getRow(rowIndex);
        if (tableRow == null) {
            return true;
        }
        for (XWPFTableCell tableCell : tableRow.getTableCells()) {
            String text = tableCell.getText();
            if (StrUtil.isEmpty(text)) {
                continue;
            }
            return false;
        }
        return true;
    }

    public static XWPFTableCell findCell(XWPFTableRow tableRow, int cellIndex) {
        XWPFTableCell tableCell = tableRow.getCell(cellIndex);
        XWPFTableCell cell = tableCell == null ? tableRow.createCell() : tableCell;
        return cell;
    }

    public static XWPFParagraph setCellValue(XWPFTableCell cell, Object value, CellStyle cellStyle) {
        int size = cell.getParagraphs().size();
        for (int i = 0; i < size; i++) {
            cell.removeParagraph(0);
        }
        XWPFParagraph xwpfParagraph = cell.addParagraph();
        XWPFRun run = xwpfParagraph.createRun();
        run.setText(StrUtil.toString(value));
        if (cellStyle != null) {
            cellStyle.set(cell, xwpfParagraph);
        }
        return xwpfParagraph;
    }

    public static XWPFParagraph setCellValue(XWPFTableRow tableRow, int pos, Object value) {
        XWPFTableCell cell = findCell(tableRow, pos);
        return setCellValue(cell, value, null);
    }

    public static CTRPr copyRunStyle(XWPFParagraph paragraph, String findContent) {
        List<XWPFRun> runs = paragraph.getRuns();
        XWPFRun originalRun = runs.get(0);
        for (XWPFRun run : runs) {
            String text = run.text();
            if (StrUtil.isEmpty(text)) {
                continue;
            }
            if (text.contains(findContent)) {
                originalRun = run;
                break;
            }
        }
        return copyRunStyle(originalRun);
    }

    public static CTRPr copyRunStyle(XWPFParagraph paragraph, int index) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) {
            return null;
        }
        XWPFRun xwpfRun = runs.get(index);
        return copyRunStyle(xwpfRun);
    }

    public static CTRPr copyRunStyle(XWPFRun originalRun) {
        if (originalRun == null) {
            return null;
        }
        CTR ctr = originalRun.getCTR();
        if (ctr == null) {
            return null;
        }
        CTRPr rPr = ctr.getRPr();
        if (rPr == null) {
            return null;
        }
        return ((CTRPr) rPr.copy());
    }

    public static void removeRunAll(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null) {
            return;
        }
        final Iterator<XWPFRun> iterator = runs.iterator();
        while (iterator.hasNext()) {
            paragraph.removeRun(0);
        }
    }

    public static CellReference setTitleInDataSheet(XSSFWorkbook workbook, String title, int column) {
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null) {
            row = sheet.createRow(0);
        }
        XSSFCell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

    public static void fullData(XWPFTable table, FullConfig fullConfig, AbsDataInput<?> fullDataInfo) throws FullException {
//        if (fullDataInfo instanceof TableFullDataInfo) {
//            if (StrUtil.isNotEmpty(fullDataInfo.getExtendKey())) {
//                Map<String, FullConfig.Position> tableStartDrawIndexMap = fullConfig.getTableStartDrawIndexMap();
//                if (!tableStartDrawIndexMap.isEmpty() && !tableStartDrawIndexMap.containsKey(fullDataInfo.getKey())) {
//                    return;
//                }
//            }
//            TableFullDataInfo tableFullDataInfo = (TableFullDataInfo) fullDataInfo;
//            tableFullDataInfo.setFullConfig(fullConfig);
//            FullContent fullContent = new FullContent();
//            fullContent.setTable(table);
//            tableFullDataInfo.full(fullContent);
//        }
    }

    public static XWPFParagraph full(AbsDataInput<?> fullDataInfo, FullContent fullContent, boolean setIndentationFirstLine) throws FullException {
        XWPFRun lastRun;
        ParagraphCreateFun paragraphCreateFun = fullContent.getParagraphCreateFun();
        XWPFParagraph paragraph = fullContent.getParagraph();
        if (fullDataInfo.isNewLine()) {
            XWPFParagraph xwpfParagraph = paragraphCreateFun.insertNewParagraph();
            lastRun = xwpfParagraph.createRun();
            paragraph = xwpfParagraph;
        } else {
            lastRun = paragraph.createRun();
        }
        int indentationFirstLine = fullDataInfo.getIndentationFirstLine();
        if (indentationFirstLine > 0 && fullDataInfo.isNotEmpty()) {
            if (setIndentationFirstLine) {
                paragraph.setIndentationFirstLine(indentationFirstLine * 105);
            }
        }
        ParagraphAlignment alignment = fullDataInfo.getAlignment();
        if (alignment != null) {
            paragraph.setAlignment(alignment);
        }
        lastRun.getCTR().setRPr(fullContent.getCtrPr());
        fullContent.setRun(lastRun);
        fullDataInfo.full(fullContent);
        return paragraph;
    }

    private static final CTString EMPTY_STRING = CTString.Factory.newInstance();

    static {
        EMPTY_STRING.setVal("");
    }

    public static CTHeight[] getCTHeight(XWPFTableRow row) {
        CTRow ctRow = row.getCtRow();
        CTTrPr trPr = ctRow.getTrPr();
        if (trPr == null) {
            return null;
        }
        return trPr.getTrHeightArray();
    }

    public static float getCTHeightIntValue(XWPFTableRow row) {
        CTHeight[] ctHeight = getCTHeight(row);
        if (ctHeight == null || ctHeight.length < 1) {
            return 0;
        }
        float r = 0F;
        for (CTHeight height : ctHeight) {
            Object val = height.getVal();
            if (val instanceof BigInteger) {
                BigInteger bigInteger = (BigInteger) val;
                r = Math.max(r, bigInteger.floatValue());
            }
        }
        return r;
    }

    public static float getCellContentRow(XWPFTableRow row, float fontSize, float cellMarginLeft, float cellMarginRight) {
        List<XWPFTableCell> tableCells = row.getTableCells();
        if (tableCells.isEmpty()) {
            return 0F;
        }
        float maxHeight = 0;
        for (XWPFTableCell tableCell : tableCells) {
            float width = tableCell.getWidth() - cellMarginLeft - cellMarginRight;
            List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
            float t = 0;
            for (XWPFParagraph paragraph : paragraphs) {
                String text = paragraph.getText();
                t += getCellContentRow(text, fontSize, width);
            }
            maxHeight = Math.max(maxHeight, t);
        }
        return maxHeight;
    }

    public static float getCellContentRow(String text, float fontSize, float width) {
        if (StrUtil.isEmpty(text)) {
            return 0;
        }
        // 像素 Points = dxa / 20
        // 英寸 Inches = Points / 72
        // 厘米 Centimeters = Inches * 2.54
        // 72pt = 25.4mm
        // 0.35277777777777777777777777777778
        // 1cm = 360000 EMUs
        float mm = TableUtil.pt2mm(fontSize);
        float dxa = TableUtil.mm2dxa(mm);
        float rowCharCount = width / dxa;
        //double rowCharCount = BigDecimal.valueOf((width / (fontSize * 4)) - 3).setScale(0, RoundingMode.UP).intValue();
        if (rowCharCount <= 0) {
            return 0;
        }
        long dataCharLength = CharUtil.dataCharLength(text);
        if (dataCharLength >= rowCharCount) {
            float max = Math.max(rowCharCount, dataCharLength);
            float min = Math.min(rowCharCount, dataCharLength);
            BigDecimal bigDecimal = BigDecimal.valueOf(max / min);
            return bigDecimal.setScale(0, RoundingMode.UP).floatValue();
        }
        return 1F;
    }

    public static float pt2dxa(float pt) {
        float mm = pt2mm(pt);
        float dxa = mm2dxa(mm);
        return dxa;
    }

    public static float pt2mm(float pt) {
        return (pt / 4F) * 0.35277777777777777777777777777778F;
    }

    public static float mm2dxa(float mm) {
        return cm2dxa(mm / 10F);
    }

    public static float cm2dxa(float cm) {
        return (cm * 1440F) / 2.54F;
    }

    public static float dxa2cm(float dxa) {
        return (dxa / 1440F) * 2.54F;
    }

    public static void main(String[] args) {
        double fontSize = 24;
        double mm = fontSize * 0.35277777777777777777777777777778;
        double cm = mm / 10f;
        double eMUs = cm * 360000;
        double dxa = eMUs / 635;
        System.out.println(dxa);
        System.out.println(((mm / 10f) * 1440) / 2.54);
    }

    public static void setCTHeight(XWPFTableRow row, int finalHeight) {
        CTRow ctRow = row.getCtRow();
        CTTrPr trPr = ctRow.getTrPr();
        if (trPr == null) {
            return;
        }
        CTHeight[] trHeightArray = trPr.getTrHeightArray();
        for (CTHeight ctHeight : trHeightArray) {
            trPr.removeTrHeight(0);
        }
        CTHeight ctHeight = trPr.addNewTrHeight();
        ctHeight.setVal(finalHeight);
    }

    public static float getFontSize(XWPFParagraph paragraph) {
        CTPPr ctpPr = paragraph.getCTPPr();
        if (ctpPr == null) {
            return 0;
        }
        CTParaRPr rPr = ctpPr.getRPr();
        if (rPr == null) {
            return 0;
        }
        CTHpsMeasure[] szArray = rPr.getSzArray();
        if (szArray == null || szArray.length < 1) {
            return 0;
        }
        return ((BigInteger) szArray[0].getVal()).intValue();
    }

    public static float getFontSize(XWPFRun xwpfRun) {
        Double fontSize = xwpfRun.getFontSizeAsDouble();
        if (fontSize == null) {
            return 0;
        }
        return fontSize.floatValue();
    }

    public static void copyHeight(XWPFTableRow targetRow, XWPFTableRow sourceRow) {
        if (targetRow == null || sourceRow == null) {
            return;
        }
        targetRow.setHeight(sourceRow.getHeight());
    }
}
