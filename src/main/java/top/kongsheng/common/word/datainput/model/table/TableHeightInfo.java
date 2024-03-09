package top.kongsheng.common.word.datainput.model.table;

import top.kongsheng.common.word.datainput.utils.TableUtil;
import lombok.Data;
import lombok.ToString;
import org.apache.poi.xwpf.usermodel.*;

import java.io.Serializable;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 表格高度信息
 *
 * @author 孔胜
 * @date 2023/9/13 10:27
 */
@Data
@ToString
public class TableHeightInfo implements Serializable {
    private final Map<XWPFTableRow, RowHeightInfo> rowHeightInfoList = new LinkedHashMap<>();
    private final Map<String, RowHeightInfo> fullKeyRowHeightInfoList = new LinkedHashMap<>();

    private float maxTotalHeight = 0F;
    private float minTotalHeight = 0F;

    private float setTotalHeight = 0F;
    private float contentTotalHeight = 0F;

    private final float cellMarginLeft;
    private final float cellMarginRight;
    private final float cellMarginBottom;
    private final float cellMarginTop;

    public TableHeightInfo(XWPFTable table) {
        this.cellMarginLeft = table.getCellMarginLeft();
        this.cellMarginRight = table.getCellMarginRight();
        this.cellMarginBottom = table.getCellMarginBottom();
        this.cellMarginTop = table.getCellMarginTop();

        this.init(table);
    }

    private void init(XWPFTable table) {
        this.update(table);
    }

    public void initFullKeyRowHeightInfo(Map<XWPFTableRow, List<String>> rowFullKeyMap) {
        this.fullKeyRowHeightInfoList.clear();
        for (Map.Entry<XWPFTableRow, List<String>> entry : rowFullKeyMap.entrySet()) {
            RowHeightInfo rowHeightInfo = this.rowHeightInfoList.get(entry.getKey());
            if (rowHeightInfo == null) {
                continue;
            }
            List<String> fullKeyList = entry.getValue();
            for (String fullKey : fullKeyList) {
                this.fullKeyRowHeightInfoList.put(fullKey, rowHeightInfo);
            }
        }
    }

    public void update(XWPFTable table) {
        final List<XWPFTableRow> rows = table.getRows();
        this.maxTotalHeight = this.minTotalHeight = this.setTotalHeight = this.contentTotalHeight = 0;
        this.rowHeightInfoList.clear();
//        float cmSum = 0F;
        for (XWPFTableRow tableRow : rows) {
            RowHeightInfo rowHeightInfo = this.findRowMaxContentHeight(tableRow);
            float setHeight = rowHeightInfo.getSetHeight();
//            cmSum += TableUtil.dxa2cm(setHeight);
            float contentHeight = rowHeightInfo.getContentHeight();
            this.maxTotalHeight += rowHeightInfo.getMaxHeight();
            this.minTotalHeight += rowHeightInfo.getMinHeight();
            this.setTotalHeight += setHeight;
            this.contentTotalHeight += contentHeight;
            this.rowHeightInfoList.put(tableRow, rowHeightInfo);
        }
//        float tableCellMarginTopBottomTotal = (cellMarginTop + cellMarginBottom) * (rows.size() - 0.5F);
//        this.setTotalHeight += tableCellMarginTopBottomTotal;
//        System.out.println("20.36 to dxa：" + TableUtil.cm2dxa(20.36F));
//        System.out.println("0.33571 to dxa：" + TableUtil.cm2dxa(0.33571F));
//        System.out.println("total cm：" + cmSum);
//        System.out.println(" cm：" + TableUtil.dxa2cm(this.setTotalHeight));
//        System.out.println("total dxa：" + this.setTotalHeight + " cm：" + TableUtil.dxa2cm(this.setTotalHeight));
//        System.out.println("------------------");
    }

    public RowHeightInfo getRowHeightInfo(String fullKey) {
        return this.fullKeyRowHeightInfoList.get(fullKey);
    }

    public RowHeightInfo getRowHeightInfo(XWPFTableRow row) {
        return this.rowHeightInfoList.get(row);
    }

    public float getCellMarginTopBottomSum() {
        return cellMarginTop + cellMarginBottom;
    }

    public float getTableSetHeight() {
        float cellMarginTopBottomSum = getCellMarginTopBottomSum();
        float tableCellMarginTopBottomTotal = cellMarginTopBottomSum * rowHeightInfoList.size();
        return this.getSetTotalHeight() + tableCellMarginTopBottomTotal;
    }

    public float getTableMaxHeight() {
        float total = 0F;
        float cellMarginTopBottomSum = getCellMarginTopBottomSum();
        for (RowHeightInfo rowHeightInfo : this.rowHeightInfoList.values()) {
            float setHeight = rowHeightInfo.getSetHeight() + cellMarginTopBottomSum;
            float contentHeight = rowHeightInfo.getContentHeight()+ cellMarginTopBottomSum;
            total += Math.max(setHeight, contentHeight) ;
        }
        //System.out.println(" tableMaxHeight(17.2)：" + TableUtil.dxa2cm(total));
        return total;
    }

    private RowHeightInfo findRowMaxContentHeight(XWPFTableRow tableRow) {
        RowHeightInfo rowHeightInfo = new RowHeightInfo();
        updateRowMaxContentHeight(rowHeightInfo, tableRow);
        return rowHeightInfo;
    }


    private void updateRowMaxContentHeight(RowHeightInfo rowHeightInfo, XWPFTableRow tableRow) {
        List<XWPFTableCell> tableCells = tableRow.getTableCells();
        float fontSize = 12;
        for (XWPFTableCell tableCell : tableCells) {
            List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                float paragraphFontSize = TableUtil.getFontSize(paragraph);
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    float fs = TableUtil.getFontSize(run);
                    if (fs == 0) {
                        fs = paragraphFontSize;
                    }
                    fontSize = Math.max(fontSize, fs);
                }
            }
        }
        float fontSizePt = fontSize * 2;
        float contentRow = TableUtil.getCellContentRow(tableRow, fontSizePt, cellMarginLeft, cellMarginRight);
//        float dxa = TableUtil.pt2dxa(fontSizePt);
        float dxa = fontSizePt * 12.15F;
//        float dxa = fontSize * 11;
        float contentHeight = (contentRow * dxa);
        float setHeight = TableUtil.getCTHeightIntValue(tableRow);
        rowHeightInfo.setContentHeight(contentHeight);
        rowHeightInfo.setSetHeight(setHeight);
    }
}
