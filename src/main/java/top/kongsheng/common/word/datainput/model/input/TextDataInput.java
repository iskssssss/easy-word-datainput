package top.kongsheng.common.word.datainput.model.input;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.date.LocalDateTimeUtil;
import cn.hutool.core.util.ObjUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.word.datainput.exception.DelRowException;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.utils.CharUtil;
import top.kongsheng.common.word.datainput.utils.FullDataInfoUtil;
import top.kongsheng.common.word.datainput.utils.TableUtil;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;
import java.util.Set;

/**
 * 填充数据
 *
 * @author 孔胜
 * @date 2023/7/28 16:03
 */
public class TextDataInput extends AbsDataInput<TextDataInput> {

    /**
     * 值
     */
    private Object value;

    private String valueStr = null;

    /**
     * 数据格式
     */
    private String dataFormat;

    /**
     * 是否加粗
     */
    private boolean bold = false;

    private boolean splitValue = false;

    private String fontColorRGB = null;

    public TextDataInput(Object value) {
        super(1);
        this.value = value;
    }

    public Object getValue() {
        return value;
    }

    public TextDataInput setValue(Object value) {
        this.valueStr = null;
        this.value = value;
        return this;
    }

    public String getDataFormat() {
        return dataFormat;
    }

    public TextDataInput setDataFormat(String dataFormat) {
        this.dataFormat = dataFormat;
        return this;
    }

    public boolean isBold() {
        return bold;
    }

    public TextDataInput setBold(boolean bold) {
        this.bold = bold;
        return this;
    }

    public boolean isSplitValue() {
        return splitValue;
    }

    public TextDataInput setSplitValue(boolean splitValue) {
        this.splitValue = splitValue;
        return this;
    }

    public String getFontColorRGB() {
        return fontColorRGB;
    }

    public TextDataInput setFontColorRGB(String fontColorRGB) {
        this.fontColorRGB = fontColorRGB;
        return this;
    }

    @Override
    public long dataLength() {
        return this.getValueStr().length();
    }

    @Override
    public long dataCharLength() {
        String valueStr = this.getValueStr();
        long dataCharLength = CharUtil.dataCharLength(valueStr);
        return dataCharLength;
    }

    @Override
    public boolean cacheAutoIndentationFirstLine(float width, float fontSize) {
        boolean autoIndentationFirstLine = isAutoIndentationFirstLine();
        if (!autoIndentationFirstLine) {
            return true;
        }
        String valueStr = this.getValueStr();
        float cellContentRow = TableUtil.getCellContentRow(valueStr, fontSize, width);
        return cellContentRow > 1F;
    }

    @Override
    public boolean isEmpty() {
        return ObjUtil.isEmpty(value);
    }

    @Override
    public List<AbsDataInput<?>> split() {
        return FullDataInfoUtil.splitValue(this);
    }

    public TextDataInput copy(String key, Object value) {
        TextDataInput result = new TextDataInput(value);
        result.setKey(key);
        result.setBold(this.bold);
        result.setIndentationFirstLine(this.getIndentationFirstLine());
        return result;
    }

    public String toValueStr() {
        if (ObjUtil.isEmpty(this.value)) {
            this.valueStr = "";
        }
        String dataFormat = this.getDataFormat();
        Object fullDataValue = this.getValue();
        if (ObjUtil.isEmpty(fullDataValue)) {
            this.valueStr = "";
        } else if (fullDataValue instanceof Date) {
            dataFormat = StrUtil.isEmpty(dataFormat) ? "yyyy-MM-dd" : dataFormat;
            this.valueStr = DateUtil.format(((Date) fullDataValue), dataFormat);
        } else if (fullDataValue instanceof LocalDateTime) {
            dataFormat = StrUtil.isEmpty(dataFormat) ? "yyyy-MM-dd" : dataFormat;
            this.valueStr = LocalDateTimeUtil.format(((LocalDateTime) fullDataValue), dataFormat);
        } else {
            this.valueStr = StrUtil.toString(fullDataValue);
        }
        return this.valueStr;
    }

    public String getValueStr() {
        if (ObjUtil.isEmpty(this.valueStr)) {
            return toValueStr();
        }
        return this.valueStr;
    }

    @Override
    public void full(FullContent fullContent) throws FullException {
        XWPFRun run = fullContent.getRun();
        boolean bold = this.isBold();
        if (bold) {
            run.setBold(true);
        }
        if (StrUtil.isNotEmpty(fontColorRGB)) {
            run.setColor(fontColorRGB);
        }
        String valueStr = this.getValueStr();
        String key = this.getKey();
        Set<String> emptyDelRowKey = this.getFullConfig().getEmptyDelRowKey();
        if (StrUtil.isEmpty(valueStr) && emptyDelRowKey.contains(key)) {
            throw new DelRowException();
        }
        run.setText(valueStr);
    }

    public static TextDataInput to(Object value) {
        return new TextDataInput(value);
    }
}
