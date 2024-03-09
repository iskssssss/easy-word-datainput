package top.kongsheng.common.word.datainput.model.input;

import cn.hutool.core.lang.Pair;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.word.datainput.model.style.DataLabelStyle;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.config.FullConfig;
import top.kongsheng.common.word.datainput.model.style.ShapeProperties;
import top.kongsheng.common.word.datainput.utils.ChartUtil;
import top.kongsheng.common.word.datainput.utils.TableUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbl;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * 图表数据
 *
 * @author 孔胜
 * @date 2023/8/1 16:09
 */
public class ChartDataInput extends AbsDataInput<ChartDataInput> {

    /**
     * 目录
     */
    private final List<String> categories;
    /**
     * 数据
     */
    private final Map<String, List<Double>> titleValueMap;

    public ChartDataInput(String key) {
        super(4, key);
        categories = new LinkedList<>();
        titleValueMap = new LinkedHashMap<>();
    }

    public void addCategories(String categories) {
        this.categories.add(categories);
    }

    public void insertCategories(int index, String categories) {
        this.categories.add(index, categories);
    }

    public void addTitleValue(String title, Double value) {
        List<Double> titleValue = this.getTitleValue(title);
        titleValue.add(value);
    }

    public void insertTitleValue(String title, int index, Double value) {
        List<Double> titleValue = this.getTitleValue(title);
        titleValue.add(index, value);
    }

    public List<Double> getTitleValue(String title) {
        return this.titleValueMap.computeIfAbsent(title, key -> new LinkedList<>());
    }

    private void setDataLabelInfo(XDDFChartData.Series series, DataLabelStyle dataLabelStyle, List<Double> sourceValueList, List<Integer> delDataLabelList) {
        CTDLbls ctdLbls = ChartUtil.findCTDLbls(series);

        FullConfig fullConfig = super.getFullConfig();
        boolean show = fullConfig.isShowDataLabel();
        if (fullConfig.getShowType() == 2) {
            if (sourceValueList.size() == delDataLabelList.size()) {
                show = false;
            } else {
                for (Integer delIndex : delDataLabelList) {
                    CTDLbl ctdLbl = ctdLbls.addNewDLbl();
                    ctdLbl.setIdx(ChartUtil.toCTUnsignedInt(delIndex));
                    ctdLbl.setDelete(ChartUtil.TRUE);
                }
            }
        }
        if (dataLabelStyle != null) {
            dataLabelStyle.paste(ctdLbls, show);
        }
    }

    @Override
    public long dataLength() {
        return titleValueMap.size();
    }

    @Override
    public boolean isEmpty() {
        return titleValueMap.isEmpty();
    }

    @Override
    public void full(FullContent fullContent) {
        XWPFChart chart = fullContent.getChart();
        chart.setAutoTitleDeleted(true);
        FullConfig fullConfig = super.getFullConfig();
        List<XDDFChartData> chartSeries = chart.getChartSeries();
        XDDFChartData xddfChartData = chartSeries.iterator().next();
        int seriesCount = xddfChartData.getSeriesCount();
        Map<Integer, Pair<ShapeProperties, DataLabelStyle>> shapePropertiesMap = new LinkedHashMap<>();
        Map<String, Integer> drawOrderMap = new LinkedHashMap<>();
        for (int i = 0, celIndex = 1; i < seriesCount; i++) {
            XDDFChartData.Series series = xddfChartData.getSeries(i);
            String seriesName = ChartUtil.findSeriesName(series);
            if (StrUtil.isNotEmpty(seriesName)) {
                drawOrderMap.put(seriesName, celIndex++);
            }
            XDDFShapeProperties shapeProperties = series.getShapeProperties();
            ShapeProperties properties = null;
            if (shapeProperties != null) {
                properties = ShapeProperties.to(shapeProperties);
            }
            CTDLbls ctdLbls = ChartUtil.findCTDLbls(series);
            DataLabelStyle dataLabelStyle = null;
            if (ctdLbls != null) {
                dataLabelStyle = new DataLabelStyle(ctdLbls);
            }
            shapePropertiesMap.put(i, new Pair<>(properties, dataLabelStyle));
        }
        if (categories.isEmpty()) {
            for (int i = 0; i < seriesCount; i++) {
                xddfChartData.removeSeries(0);
            }
            return;
        }
        int numOfPoints = categories.size();
        int drawRowNum = numOfPoints;
        List<String> tCategories = categories;
        int showDataNum = fullConfig.getShowDataNum();
        if (showDataNum > -1) {
            drawRowNum = Math.min(showDataNum, numOfPoints);
            if (showDataNum != numOfPoints) {
                tCategories = new LinkedList<>();
                for (int i = 0; i < drawRowNum; i++) {
                    tCategories.add(categories.get(i));
                }
            }
        }
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, drawRowNum, 0, 0));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(tCategories.toArray(new String[]{}), categoryDataRange, 0);
        int colIndex = drawOrderMap.size() <= 0 ? 1 : drawOrderMap.size();
        for (Map.Entry<String, List<Double>> entry : this.titleValueMap.entrySet()) {
            colIndex++;
            String key = entry.getKey();
            if (drawOrderMap.containsKey(key)) {
                continue;
            }
            drawOrderMap.put(key, colIndex);
        }
        for (Map.Entry<String, Integer> entry : drawOrderMap.entrySet()) {
            String titleStr = entry.getKey();
            Integer cellIndex = entry.getValue();
            List<Double> values = this.titleValueMap.get(titleStr);
            if (values == null) {
                continue;
            }
            List<Integer> delDataLabelList = new LinkedList<>();
            Double[] elements = new Double[drawRowNum];
            for (int i = 0; i < drawRowNum; i++) {
                elements[i] = values.get(i);
                if (elements[i] == 0) {
                    delDataLabelList.add(i);
                }
            }
            String valuesDataRange = chart.formatRange(new CellRangeAddress(1, drawRowNum, cellIndex, cellIndex));
            XDDFNumericalDataSource<Double> valuesData = XDDFDataSourcesFactory.fromArray(elements, valuesDataRange, cellIndex);
            XDDFChartData.Series series = xddfChartData.addSeries(categoriesData, valuesData);
            Pair<ShapeProperties, DataLabelStyle> stylePair = shapePropertiesMap.get(cellIndex - 1);
            try {
                series.setTitle(titleStr, TableUtil.setTitleInDataSheet(chart.getWorkbook(), titleStr, cellIndex));
            } catch (IOException | InvalidFormatException e) {
                throw new RuntimeException(e);
            }
            if (stylePair == null) {
                continue;
            }
            ShapeProperties shapeProperties = stylePair.getKey();
            DataLabelStyle dataLabelStyle = stylePair.getValue();
            if (shapeProperties != null) {
                series.setShapeProperties(shapeProperties.copy());
            }
            // 设置图表数据标签信息
            this.setDataLabelInfo(series, dataLabelStyle, values, delDataLabelList);
        }
        for (int i = 0; i < seriesCount; i++) {
            xddfChartData.removeSeries(0);
        }
        chart.plot(xddfChartData);
    }
}
