package top.kongsheng.common.word.datainput.model.input;

import top.kongsheng.common.word.datainput.core.FullDataMap;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.config.FullConfig;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * 填充数据 抽象类
 *
 * @author 孔胜
 * @date 2023/8/2 17:19
 */
public abstract class AbsDataInput<T> implements Serializable {
    private static final long serialVersionUID = 42L;

    /**
     * 填充类型：1.文本 2.图片 3.表格 4.图表
     */
    private final int fullType;

    /**
     * 标识
     */
    private String key;
    /**
     * 扩展标识
     */
    private String extendKey;

    /**
     * 首行缩进
     */
    private int indentationFirstLine = 0;

    /**
     * 自动判断是否需要首行缩进
     */
    private boolean autoIndentationFirstLine = false;

    /**
     * 是否换行
     */
    private boolean newLine = false;

    private ParagraphAlignment alignment;

    private FullConfig fullConfig;

    /**
     * 填充参数
     */
    protected FullDataMap fullDataMap;

    protected AbsDataInput(int fullType) {
        this.fullType = fullType;
    }

    protected AbsDataInput(int fullType, String key) {
        this.fullType = fullType;
        this.key = key;
    }

    public String getKey() {
        return key;
    }

    public T setKey(String key) {
        this.key = key;
        return ((T) this);
    }

    public int getFullType() {
        return fullType;
    }

    public String getExtendKey() {
        return extendKey;
    }

    public void setExtendKey(String extendKey) {
        this.extendKey = extendKey;
    }

    public int getIndentationFirstLine() {
        return indentationFirstLine;
    }

    public T setIndentationFirstLine(int indentationFirstLine) {
        this.indentationFirstLine = indentationFirstLine;
        return ((T) this);
    }

    public boolean isAutoIndentationFirstLine() {
        return autoIndentationFirstLine;
    }

    public T setAutoIndentationFirstLine(boolean autoIndentationFirstLine) {
        this.autoIndentationFirstLine = autoIndentationFirstLine;
        return ((T) this);
    }

    public boolean isNewLine() {
        return newLine;
    }

    public T setNewLine(boolean newLine) {
        this.newLine = newLine;
        return ((T) this);
    }

    public ParagraphAlignment getAlignment() {
        return alignment;
    }

    public T setAlignment(ParagraphAlignment alignment) {
        this.alignment = alignment;
        return ((T) this);
    }

    public FullConfig getFullConfig() {
        return fullConfig;
    }

    public T setFullConfig(FullConfig fullConfig) {
        this.fullConfig = fullConfig;
        return ((T) this);
    }

    public FullDataMap getFullDataMap() {
        return fullDataMap;
    }

    public T setFullDataMap(FullDataMap fullDataMap) {
        this.fullDataMap = fullDataMap;
        return ((T) this);
    }

    /**
     * 校验是否自动换行
     *
     * @param width    内容所在位置容器宽度
     * @param fontSize 字体大小
     * @return 是否自动换行
     */
    public boolean cacheAutoIndentationFirstLine(float width, float fontSize) {
        return true;
    }

    /**
     * 获取数据长度
     *
     * @return 数据长度
     */
    public abstract long dataLength();

    /**
     * 获取数据字符长度
     *
     * @return 数据字符长度
     */
    public long dataCharLength() {
        return dataLength();
    }

    /**
     * 填充内容是否为空
     *
     * @return 是否为空
     */
    public abstract boolean isEmpty();

    /**
     * 填充内容是否不为空
     *
     * @return 是否不为空
     */
    public boolean isNotEmpty() {
        return !isEmpty();
    }

    public List<AbsDataInput<?>> split() {
        ArrayList<AbsDataInput<?>> absDataInputs = new ArrayList<>();
        absDataInputs.add(this);
        return absDataInputs;
    }

    /**
     * 数据填充
     *
     * @param fullContent 填充参数
     * @throws FullException 填充异常
     */
    public abstract void full(FullContent fullContent) throws FullException;
}
