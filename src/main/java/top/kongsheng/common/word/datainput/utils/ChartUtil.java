package top.kongsheng.common.word.datainput.utils;

import cn.hutool.core.util.StrUtil;
import cn.hutool.json.JSONUtil;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.config.FullConfig;
import top.kongsheng.common.word.datainput.model.input.AbsDataInput;
import top.kongsheng.common.word.datainput.model.input.ChartDataInput;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.text.XDDFTextBody;
import org.apache.poi.xddf.usermodel.text.XDDFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.impl.CTSerTxImpl;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * 图表工具类
 *
 * @author 孔胜
 * @date 2023/8/1 16:02
 */
public class ChartUtil {

    public static Map<String, XWPFChart> findChart(XWPFDocument document) {
        List<XWPFChart> charts = document.getCharts();
        Map<String, XWPFChart> chartMap = new LinkedHashMap<>();
        for (XWPFChart chart : charts) {
            XDDFTitle title = chart.getTitle();
            if (title == null) {
                continue;
            }
            XDDFTextBody body = title.getBody();
            StringBuilder titleSB = new StringBuilder();
            List<XDDFTextParagraph> paragraphs = body.getParagraphs();
            for (XDDFTextParagraph paragraph : paragraphs) {
                titleSB.append(paragraph.getText());
            }
            chartMap.put(titleSB.toString(), chart);
        }
        return chartMap;
    }

    public static void fullData(XWPFChart chart, FullConfig fullConfig, AbsDataInput<?> fullDataInfo) throws FullException {
        if (fullDataInfo instanceof ChartDataInput) {
            ChartDataInput chartDataInput = (ChartDataInput) fullDataInfo;
            chartDataInput.setFullConfig(fullConfig);
            FullContent fullContent = new FullContent(-1, -1);
            fullContent.setChart(chart);
            chartDataInput.full(fullContent);
        }
    }

    public static CTDLbls findCTDLbls(XDDFChartData.Series series) {
        CTDLbls dLbls = null;
        if (series instanceof XDDFLineChartData.Series) {
            XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) series;
            CTLineSer ctLineSer = series1.getCTLineSer();
            dLbls = ctLineSer.getDLbls();
            if (dLbls == null) {
                dLbls = ctLineSer.addNewDLbls();
            }
        } else if (series instanceof XDDFBarChartData.Series) {
            XDDFBarChartData.Series series1 = (XDDFBarChartData.Series) series;
            CTBarSer ctBarSer = series1.getCTBarSer();
            dLbls = ctBarSer.getDLbls();
            if (dLbls == null) {
                dLbls = ctBarSer.addNewDLbls();
            }
        } else if (series instanceof XDDFDoughnutChartData.Series) {
            XDDFDoughnutChartData.Series series1 = (XDDFDoughnutChartData.Series) series;
            CTPieSer ctPieSer = series1.getCTPieSer();
            dLbls = ctPieSer.getDLbls();
            if (dLbls == null) {
                dLbls = ctPieSer.addNewDLbls();
            }
        }
        return dLbls;
    }

    public static CTSerTx findCTSerTx(XDDFChartData.Series series) {
        try {
            Object ctStrRef = GET_SERIES_TEXT.invoke(series);
            if (ctStrRef instanceof CTSerTxImpl) {
                CTSerTxImpl ctStrRef1 = (CTSerTxImpl) ctStrRef;
                return ctStrRef1;
            }
            return null;
        } catch (IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static String findSeriesName(XDDFChartData.Series series) {
        CTSerTx ctSerTx = findCTSerTx(series);
        if (ctSerTx == null) {
            return null;
        }
        CTStrRef strRef = ctSerTx.getStrRef();
        if (strRef == null) {
            return null;
        }
        CTStrData strCache = strRef.getStrCache();
        if (strCache == null) {
            return null;
        }
        CTStrVal[] ptArray = strCache.getPtArray();
        if (ptArray == null || ptArray.length < 1) {
            return null;
        }
        return ptArray[0].getV();
    }

    public static CTBoolean toCTBoolean(boolean bool) {
        CTBoolean ctBoolean = CTBoolean.Factory.newInstance();
        ctBoolean.setVal(bool);
        return ctBoolean;
    }

    public static CTUnsignedInt toCTUnsignedInt(long val) {
        CTUnsignedInt cTUnsignedInt = CTUnsignedInt.Factory.newInstance();
        cTUnsignedInt.setVal(val);
        return cTUnsignedInt;
    }

    public static final CTBoolean TRUE = toCTBoolean(true);
    public static final CTBoolean FALSE = toCTBoolean(false);

    private final static Class<XDDFChartData.Series> seriesClass = XDDFChartData.Series.class;
    private static Method GET_SERIES_TEXT;

    static {
        try {
            GET_SERIES_TEXT = seriesClass.getDeclaredMethod("getSeriesText");
            GET_SERIES_TEXT.setAccessible(true);
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
    }

    public static FullConfig getFullConfig(XWPFChart xwpfChart) {
        try {
            XSSFWorkbook workbook = xwpfChart.getWorkbook();
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row = sheet.getRow(0);
            if (row == null) {
                return FullConfig.empty();
            }
            XSSFCell cell = row.getCell(0);
            if (cell == null) {
                return FullConfig.empty();
            }
            String cellValue = cell.getStringCellValue();
            if (StrUtil.isEmpty(cellValue)) {
                return FullConfig.empty();
            }
            cellValue = StrUtil.trim(cellValue);
            if (cellValue.startsWith("{")) {
                cell.setCellValue(" ");
                return Optional.ofNullable(JSONUtil.toBean(cellValue, FullConfig.class)).orElse(FullConfig.empty());
            }
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
        }
        return FullConfig.empty();
    }
}
