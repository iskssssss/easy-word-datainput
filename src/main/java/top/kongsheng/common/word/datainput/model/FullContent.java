package top.kongsheng.common.word.datainput.model;

import lombok.Data;
import lombok.ToString;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

import java.io.Serializable;

/**
 * 填充内容
 *
 * @author 孔胜
 * @date 2023/9/7 9:59
 */
@Data
@ToString
public class FullContent implements Serializable {
    private static final long serialVersionUID = 42L;
    /**
     * 图表
     */
    private XWPFChart chart;
    /**
     * 表格
     */
    private XWPFTable table;
    /**
     * 段落
     */
    private XWPFParagraph paragraph;
    /**
     * 文本区域
     */
    private XWPFRun run;
    /**
     * 段落创建器
     */
    private ParagraphCreateFun paragraphCreateFun;
    /**
     * 样式
     */
    private CTRPr ctrPr;
    /**
     * 宽度
     */
    private float width;
    /**
     * 字体大小
     */
    private float fontSize;

    private int rowIndex;

    private int cellIndex;

    public FullContent(int rowIndex, int cellIndex) {
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
    }
}
