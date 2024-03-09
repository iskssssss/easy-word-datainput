package top.kongsheng.common.word.datainput.model.style;

import lombok.Getter;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;

/**
 * 图表样式
 *
 * @author 孔胜
 * @date 2023/8/2 17:44
 */
@Getter
public class ShapeProperties {
    private final XDDFFillProperties fillProperties;
    private final XDDFLineProperties lineProperties;
    private final XDDFShape3D shape3D;
    private final CTShapeProperties xmlObject;
    private final BlackWhiteMode blackWhiteMode;
    private final XDDFTransform2D transform2D;
    private final XDDFScene3D scene3D;
    private final XDDFPresetGeometry2D presetGeometry2D;
    private final XDDFExtensionList extensionList;
    private final XDDFEffectList effectList;
    private final XDDFEffectContainer effectContainer;
    private final XDDFCustomGeometry2D customGeometry2D;

    public ShapeProperties(XDDFChartData.Series series) {
        this(series.getShapeProperties());
    }

    public ShapeProperties(XDDFShapeProperties shapeProperties) {
        this.fillProperties = shapeProperties.getFillProperties();
        this.lineProperties = shapeProperties.getLineProperties();
        this.shape3D = shapeProperties.getShape3D();
        this.xmlObject = shapeProperties.getXmlObject();
        this.blackWhiteMode = shapeProperties.getBlackWhiteMode();
        this.transform2D = shapeProperties.getTransform2D();
        this.scene3D = shapeProperties.getScene3D();
        this.presetGeometry2D = shapeProperties.getPresetGeometry2D();
        this.extensionList = shapeProperties.getExtensionList();
        this.effectList = shapeProperties.getEffectList();
        this.effectContainer = shapeProperties.getEffectContainer();
        this.customGeometry2D = shapeProperties.getCustomGeometry2D();
    }

    public XDDFShapeProperties copy() {
        XDDFShapeProperties result = new XDDFShapeProperties();
        result.setFillProperties(this.fillProperties);
        result.setLineProperties(this.lineProperties);
        result.setShape3D(shape3D);
        result.setBlackWhiteMode(blackWhiteMode);
        result.setTransform2D(transform2D);
        result.setScene3D(scene3D);
        result.setPresetGeometry2D(presetGeometry2D);
        result.setExtensionList(extensionList);
        result.setEffectList(effectList);
        result.setEffectContainer(effectContainer);
        result.setCustomGeometry2D(customGeometry2D);
        return result;
    }

    public static ShapeProperties to(XDDFShapeProperties shapeProperties) {
        return new ShapeProperties(shapeProperties);
    }
}
