package top.kongsheng.common.word.datainput.model.style;

import top.kongsheng.common.word.datainput.utils.ChartUtil;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;

import java.util.function.Consumer;
import java.util.function.Supplier;

/**
 * 数据标签样式
 *
 * @author 孔胜
 * @date 2023/8/14 14:29
 */
public class DataLabelStyle {

    private final CTShapeProperties shapeProperties;
    private final CTTextBody textBody;
    /**
     * 设置是否显示数据点的值（value）。值为 true 表示显示，值为 false 表示不显示。
     */
    private final boolean showVal;
    /**
     * 设置是否显示数据点的分类名称（category name）。值为 true 表示显示，值为 false 表示不显示。
     */
    private final boolean showCatName;
    /**
     * 设置是否显示数据点的图例键（legend key）。值为 true 表示显示，值为 false 表示不显示。
     */
    private final boolean showLegendKey;
    /**
     * 设置是否显示数据点的百分比。值为 true 表示显示，值为 false 表示不显示。
     */
    private final boolean showPercent;
    /**
     * 设置是否显示数据点的气泡大小（bubble size）。值为 true 表示显示，值为 false 表示不显示。
     */
    private final boolean showBubbleSize;
    /**
     * 设置是否显示数据点的引导线（leader lines）。值为 true 表示显示，值为 false 表示不显示。
     */
    private final boolean showLeaderLines;
    /**
     * 设置是否显示数据点的系列名称（series name）。值为 true 表示显示，值为 false 表示不显示。
     */
    private final boolean showSerName;

    public DataLabelStyle(CTDLbls ctdLbls) {
        CTShapeProperties spPr = ctdLbls.getSpPr();
        this.shapeProperties = spPr == null ? null : ((CTShapeProperties) spPr.copy());
        CTTextBody txPr = ctdLbls.getTxPr();
        this.textBody = txPr == null ? null : ((CTTextBody) txPr.copy());
        this.showVal = get(ctdLbls::getShowVal);
        this.showCatName = get(ctdLbls::getShowCatName);
        this.showLegendKey = get(ctdLbls::getShowLegendKey);
        this.showPercent = get(ctdLbls::getShowPercent);
        this.showBubbleSize = get(ctdLbls::getShowBubbleSize);
        this.showLeaderLines = get(ctdLbls::getShowLeaderLines);
        this.showSerName = get(ctdLbls::getShowSerName);
    }

    public void paste(CTDLbls ctdLbls, boolean showDataLabel) {
        if (this.shapeProperties != null) {
            ctdLbls.setSpPr(this.shapeProperties);
        }
        if (this.textBody != null) {
            ctdLbls.setTxPr(this.textBody);
        }
        this.set(ctdLbls::setShowVal, this.showVal && showDataLabel);
        this.set(ctdLbls::setShowCatName, this.showCatName && showDataLabel);
        this.set(ctdLbls::setShowLegendKey, this.showLegendKey && showDataLabel);
        this.set(ctdLbls::setShowPercent, this.showPercent && showDataLabel);
        this.set(ctdLbls::setShowBubbleSize, this.showBubbleSize && showDataLabel);
        this.set(ctdLbls::setShowLeaderLines, this.showLeaderLines && showDataLabel);
        this.set(ctdLbls::setShowSerName, this.showSerName && showDataLabel);
    }

    private boolean get(Supplier<CTBoolean> get) {
        CTBoolean ctBoolean = get.get();
        if (ctBoolean == null) {
            return false;
        }
        return ctBoolean.getVal();
    }

    private void set(Consumer<CTBoolean> setConsumer, boolean value) {
        setConsumer.accept(ChartUtil.toCTBoolean(value));
    }

    public CTShapeProperties getShapeProperties() {
        return shapeProperties;
    }

    public CTTextBody getTextBody() {
        return textBody;
    }

    public boolean isShowVal() {
        return showVal;
    }

    public boolean isShowCatName() {
        return showCatName;
    }

    public boolean isShowLegendKey() {
        return showLegendKey;
    }

    public boolean isShowPercent() {
        return showPercent;
    }

    public boolean isShowBubbleSize() {
        return showBubbleSize;
    }

    public boolean isShowLeaderLines() {
        return showLeaderLines;
    }

    public boolean isShowSerName() {
        return showSerName;
    }
}
