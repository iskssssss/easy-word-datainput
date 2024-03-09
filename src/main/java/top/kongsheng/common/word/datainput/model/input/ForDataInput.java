package top.kongsheng.common.word.datainput.model.input;

import cn.hutool.core.util.ReflectUtil;
import top.kongsheng.common.word.datainput.model.style.CellStyle;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.config.FullConfig;
import top.kongsheng.common.word.datainput.utils.TableUtil;
import lombok.Data;
import lombok.ToString;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.Serializable;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

/**
 * 循环填充
 *
 * @author 孔胜
 * @date 2024/1/10 11:34
 */
public class ForDataInput extends AbsDataInput<ForDataInput> {

    private List<?> dataList;
    private Map<Integer, String> titleIndexMap;
    private ForInfo forInfo;
    private int maxRowNum = 0;

    public ForDataInput(String key) {
        super(3, key);
    }

    @Override
    public long dataLength() {
        return dataList.size();
    }

    @Override
    public boolean isEmpty() {
        return dataList.isEmpty();
    }

    @Override
    public void full(FullContent fullContent) {
        FullConfig fullConfig = getFullConfig();
        XWPFTable table = fullContent.getTable();
        int rowIndex = fullContent.getRowIndex();
        Map<Integer, CellStyle> cellStyleMap = new LinkedHashMap<>();
        Set<Map.Entry<Integer, String>> titleIndexMapEntries = this.titleIndexMap.entrySet();
        int dataListSize = this.dataList.size();
        if (dataListSize == 0 && this.maxRowNum <= 0) {
            this.maxRowNum = 1;
        }
        ForInfo forInfo = this.forInfo;
        int startIndex = forInfo.getStartIndex(), endIndex = forInfo.getEndIndex(), step = forInfo.getStep();
        int maxRowNum = this.maxRowNum > 0 ? this.maxRowNum : dataListSize;
        if (endIndex == -1) {
            endIndex = maxRowNum;
        }
        for (int i = startIndex, ri = rowIndex; i < maxRowNum && i < endIndex; i += step, ri++) {
            boolean overIndex = i >= dataListSize;
            Object obj = overIndex ? null : this.dataList.get(i);
            for (Map.Entry<Integer, String> entry : titleIndexMapEntries) {
                Integer ci = entry.getKey();
                String title = entry.getValue();
                Object fieldValue = overIndex ? "" : ("INDEX".equals(title) ? (i + 1) : ReflectUtil.getFieldValue(obj, title));
                XWPFTableRow tableRow;
                if (ri == rowIndex || ci > 0 || fullConfig.getTableFullType() == 1) {
                    tableRow = table.getRow(ri);
                    if (tableRow == null) {
                        tableRow = table.createRow();
                        TableUtil.copyHeight(tableRow, table.getRow(ri - 1));
                    }
                } else {
                    tableRow = table.insertNewTableRow(ri);
                }
                XWPFTableCell cell = TableUtil.findCell(tableRow, ci);
                XWPFTableRow finalTableRow = tableRow;
                CellStyle cellStyle = cellStyleMap.computeIfAbsent(ci, k -> new CellStyle(finalTableRow.getCell(k)));
                TableUtil.setCellValue(cell, fieldValue, cellStyle);
            }
        }
        this.merge(table, fullConfig);
    }

    public void setDataList(List<?> dataList) {
        this.dataList = dataList;
    }

    public void setTitleIndexMap(Map<Integer, String> titleIndexMap) {
        this.titleIndexMap = titleIndexMap;
    }

    public void setForInfo(ForInfo forInfo) {
        this.forInfo = forInfo;
    }

    public void setMaxRowNum(int maxRowNum) {
        this.maxRowNum = maxRowNum;
    }

    protected void merge(XWPFTable table, FullConfig fullConfig) {
        List<XWPFTableRow> rows = table.getRows();
        int rowsSize = rows.size();
        List<FullConfig.MergeCellInfo> autoMergeCells = fullConfig.getAutoMergeCells();
        for (FullConfig.MergeCellInfo autoMergeCell : autoMergeCells) {
            if (autoMergeCell.isEmpty()) {
                continue;
            }
            int sx = autoMergeCell.getSx(), sy = autoMergeCell.getSy();
            sy = sy == -1 ? rowsSize : sy;
            int ex = autoMergeCell.getEx(), ey = autoMergeCell.getEy();
            ey = ey == -1 ? rowsSize : ey;
            mergeCellsVertically(table, sx, sy, ex, ey);
        }
    }

    private static void mergeCellsVertically(XWPFTable table, int sx, int sy, int ex, int ey) {
        List<XWPFTableRow> rows = table.getRows();
        int rowsSize = rows.size();
        for (int y = sy; y < rowsSize && y <= ey; y++) {
            XWPFTableRow row = table.getRow(y);
            List<XWPFTableCell> cells = row.getTableCells();
            int cellsSize = cells.size();
            for (int x = sx; x < cellsSize && x <= ex; x++) {
                XWPFTableCell cell = row.getCell(x);
                CTVMerge ctvMerge = cell.getCTTc().addNewTcPr().addNewVMerge();
                if (y == sy) {
                    ctvMerge.setVal(STMerge.RESTART);
                    continue;
                }
                ctvMerge.setVal(STMerge.CONTINUE);
            }
        }
    }

    public static final String PREFIX = "${F ";
    public static final int PREFIX_LENGTH = PREFIX.length();
    public static final Pattern FIELD_PATTERN = Pattern.compile("%(.*)%");

    public static boolean check(String sourceKey) {
        return sourceKey.startsWith(PREFIX);
    }

    public static ForInfo parseForInfo(String sourceKey) {
        String keyFormat = sourceKey.substring(PREFIX_LENGTH, sourceKey.indexOf(":"));
        String[] splits = keyFormat.split(",");
        String key = splits[0];
        int startIndex = 0, endIndex = -1, step = 1;
        if (splits.length > 1) {
            startIndex = Integer.parseInt(splits[1]);
            endIndex = Integer.parseInt(splits[2]);
            step = Integer.parseInt(splits[3]);
        }
        return new ForInfo(key, startIndex, endIndex, step);
    }

    @Data
    @ToString
    public static class ForInfo implements Serializable {
        private String key;
        private int startIndex;
        private int endIndex;
        private int step;

        public ForInfo(String key, int startIndex, int endIndex, int step) {
            this.key = key;
            this.startIndex = startIndex;
            this.endIndex = endIndex;
            this.step = step;
        }
    }
}
