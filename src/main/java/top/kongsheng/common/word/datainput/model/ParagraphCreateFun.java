package top.kongsheng.common.word.datainput.model;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlCursor;

import java.util.function.Function;
import java.util.function.Supplier;

/**
 * 段落添加方法
 *
 * @author 孔胜
 * @date 2023/8/23 13:59
 */
public class ParagraphCreateFun {
    private final Function<XmlCursor, XWPFParagraph> insertNewParagraphFun;
    private final Supplier<XWPFParagraph> addParagraphFun;
    private XmlCursor xmlCursor;

    public ParagraphCreateFun(Function<XmlCursor, XWPFParagraph> insertNewParagraphFun, Supplier<XWPFParagraph> addParagraphFun) {
        this.insertNewParagraphFun = insertNewParagraphFun;
        this.addParagraphFun = addParagraphFun;
    }

    public XWPFParagraph insertNewParagraph(XmlCursor xmlCursor) {
        return insertNewParagraphFun.apply(xmlCursor);
    }

    public XWPFParagraph insertNewParagraph() {
        if (this.xmlCursor == null) {
            return this.addParagraph();
        }
        return insertNewParagraphFun.apply(this.xmlCursor);
    }

    public XWPFParagraph addParagraph() {
        return addParagraphFun.get();
    }

    public void setXmlCursor(XmlCursor xmlCursor) {
        this.xmlCursor = xmlCursor;
    }
}
