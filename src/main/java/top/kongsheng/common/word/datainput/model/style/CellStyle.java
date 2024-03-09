package top.kongsheng.common.word.datainput.model.style;

import top.kongsheng.common.word.datainput.utils.TableUtil;
import lombok.Getter;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.List;
import java.util.Optional;

/**
 * 单元格样式
 *
 * @author 孔胜
 * @date 2023/8/2 17:47
 */
@Getter
public class CellStyle {
    private final XWPFTableCell.XWPFVertAlign verticalAlignment;
    private final CTPPr ctpPr;
    private final CTRPr ctrPr;
    private final CTTcBorders tcBorders;

    public CellStyle(XWPFTableCell tableCell) {
        this.verticalAlignment = tableCell.getVerticalAlignment();
        XWPFParagraph paragraph = tableCell.getParagraphs().iterator().next();
        this.ctpPr = ((CTPPr) paragraph.getCTPPr().copy());
        this.ctrPr = TableUtil.copyRunStyle(paragraph, 0);
        CTTcPr tcPr = tableCell.getCTTc().getTcPr();
        if (tcPr == null) {
            this.tcBorders = null;
            return;
        }
        CTTcBorders tcBorders = tcPr.getTcBorders();
        if (tcBorders == null) {
            this.tcBorders = null;
            return;
        }
        this.tcBorders = ((CTTcBorders) tcBorders.copy());
    }

    public void set(XWPFTableCell cell, XWPFParagraph paragraph) {
        if (this.tcBorders != null) {
            CTTc ctTc = cell.getCTTc();
            CTTcPr tcPr = Optional.ofNullable(ctTc.getTcPr()).orElse(ctTc.addNewTcPr());
            tcPr.setTcBorders(this.tcBorders);
        }
        cell.setVerticalAlignment(this.getVerticalAlignment());
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.getCTPPr().setJc(this.ctpPr.getJc());
        CTParaRPr ctpPrRPr = this.ctpPr.getRPr();
        paragraph.getCTPPr().setRPr(ctpPrRPr);
        List<XWPFRun> runs = paragraph.getRuns();
        for (XWPFRun run : runs) {
            CTR ctr = run.getCTR();
            if (ctrPr != null) {
                ctr.setRPr(ctrPr);
                continue;
            }
            if (ctpPrRPr == null) {
                continue;
            }
            CTRPr rPr = ctr.getRPr() == null ? ctr.addNewRPr() : ctr.getRPr();
            rPr.setRFontsArray(ctpPrRPr.getRFontsArray());
            rPr.setSzArray(ctpPrRPr.getSzArray());
            rPr.setSzCsArray(ctpPrRPr.getSzCsArray());
        }
    }

//  <w:p w14:paraId="3C4CFAB4" w14:textId="77777777" w:rsidR="008B7108" w:rsidRPr="00F40C38" w:rsidRDefault="00000000">
//    <w:pPr>
//      <w:jc w:val="center"/>
//      <w:rPr>
//        <w:rFonts w:ascii="仿宋" w:eastAsia="仿宋" w:hAnsi="仿宋"/>
//        <w:b/>
//        <w:bCs/>
//        <w:sz w:val="28"/>
//        <w:szCs w:val="28"/>
//      </w:rPr>
//    </w:pPr>
//    <w:r w:rsidRPr="00F40C38">
//      <w:rPr>
//        <w:rFonts w:ascii="仿宋" w:eastAsia="仿宋" w:hAnsi="仿宋" w:hint="eastAsia"/>
//        <w:b/>
//        <w:bCs/>
//        <w:sz w:val="28"/>
//        <w:szCs w:val="28"/>
//      </w:rPr>
//      <w:lastRenderedPageBreak/>
//      <w:t>外单位</w:t>
//    </w:r>
//  </w:p>

    //     <w:pPr>
    //     <w:jc w:val="center"/>
    //     <w:rPr>
    //     <w:rFonts w:ascii="仿宋" w:eastAsia="仿宋" w:hAnsi="仿宋"/>
    //     <w:sz w:val="28"/>
    //     <w:szCs w:val="28"/>
    //     </w:rPr>
    //     </w:pPr>
}
