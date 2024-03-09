package top.kongsheng.common.word.datainput.core;

import cn.hutool.core.date.StopWatch;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.util.ObjUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.word.datainput.convert.InputDataConvert;
import top.kongsheng.common.word.datainput.exception.DelRowException;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.model.ParagraphCreateFun;
import top.kongsheng.common.word.datainput.model.RunInfo;
import top.kongsheng.common.word.datainput.config.FullConfig;
import top.kongsheng.common.word.datainput.extend.ExtendDataInputUtil;
import top.kongsheng.common.word.datainput.model.TemplateInfo;
import top.kongsheng.common.word.datainput.model.input.AbsDataInput;
import top.kongsheng.common.word.datainput.model.input.BatchDataInput;
import top.kongsheng.common.word.datainput.model.input.ForDataInput;
import top.kongsheng.common.word.datainput.model.input.TextDataInput;
import top.kongsheng.common.word.datainput.model.table.RowHeightInfo;
import top.kongsheng.common.word.datainput.model.table.TableHeightInfo;
import top.kongsheng.common.word.datainput.utils.ChartUtil;
import top.kongsheng.common.word.datainput.utils.TableUtil;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.ooxml.util.PackageHelper;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.stream.Collectors;

import static top.kongsheng.common.word.datainput.WordDataInputManage.WORD_DATA_INPUT_FILE_SAVE_DIR;

/**
 * word数据填充
 *
 * @author 孔胜
 * @date 2023/7/20 16:42
 */
@Log4j2
public final class WordDataInputCore implements Closeable {

    /**
     * 结果文件保存路径
     */
    private final String targetFilePath;
    /**
     * 结果文件名称
     */
    private final String targetFileName;
    /**
     * 结果文件保存全路径
     */
    private final String targetFileFullPath;
    /**
     * 文档信息
     */
    private final XWPFDocument document;
    /**
     * 填充模板信息
     */
    private final TemplateInfo templateInfo;
    /**
     * 填充数据转换器
     */
    private InputDataConvert inputDataConvert;
    /**
     * 填充参数
     */
    protected final FullDataMap fullDataMap = new FullDataMap();

    private BigInteger szH;

    private BigInteger top;

    private BigInteger bottom;

    public WordDataInputCore(TemplateInfo templateInfo) throws FullException {
        this(templateInfo, WORD_DATA_INPUT_FILE_SAVE_DIR);
    }

    public WordDataInputCore(TemplateInfo templateInfo, String targetFilePath) throws FullException {
        this.templateInfo = templateInfo;
        this.targetFilePath = targetFilePath;
        this.targetFileName = "RESULT_" + System.currentTimeMillis();
        this.targetFileFullPath = targetFilePath + File.separator + this.targetFileName + "." + this.templateInfo.getFileSuffix();
        // 加载模板文件
        try {
            this.document = new XWPFDocument(PackageHelper.open(this.templateInfo.createInputStream(), true));
            CTBody body = this.document.getDocument().getBody();
            if (body != null) {
                CTSectPr sectPr = body.getSectPr();
                if (sectPr != null) {
                    CTPageSz pgSz = sectPr.getPgSz();
                    if (pgSz != null) {
                        Object szH = pgSz.getH();
                        if (szH instanceof BigInteger) {
                            this.szH = ((BigInteger) szH);
                        }
                    }
                    CTPageMar pageMar = sectPr.getPgMar();
                    if (pageMar != null) {
                        this.top = ((BigInteger) pageMar.getTop());
                        this.bottom = ((BigInteger) pageMar.getBottom());
                    }
                }
            }
        } catch (IOException e) {
            throw new FullException("模板文件加载失败。", e);
        }
    }

    /**
     * 设置填充数据转换器
     *
     * @param inputDataConvert 填充数据转换器
     */
    public void setInputDataConvert(InputDataConvert inputDataConvert) {
        this.inputDataConvert = inputDataConvert;
    }

    public void run() throws FullException {
        try {
            this.fill();
        } catch (FullException e) {
            throw e;
        } catch (Exception e) {
            throw new FullException("数据填充失败", e);
        }
    }

    /**
     * 处理变量表
     */
    private void handleVariableTable() throws FullException {
        List<IBodyElement> bodyElements = this.document.getBodyElements();
        int size = bodyElements.size();
        for (int i = size - 1; i >= 0; i--) {
            IBodyElement element = bodyElements.get(i);
            if (!(element instanceof XWPFTable)) {
                continue;
            }
            XWPFTable table = (XWPFTable) element;
            String tableName = TableUtil.getTableName(table);
            if (!"$变量表".equals(tableName)) {
                continue;
            }
            List<XWPFTableRow> rowList = table.getRows();
            for (XWPFTableRow row : rowList) {
                List<XWPFTableCell> tableCellList = row.getTableCells();
                if (tableCellList.size() < 2) {
                    continue;
                }
                String key = tableCellList.get(0).getText().trim();
                if (StrUtil.isEmpty(key)) {
                    continue;
                }
                XWPFTableCell tableCell = tableCellList.get(1);
                List<RunInfo> runInfoList = new LinkedList<>();
                for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                    runInfoList.addAll(getRunInfo(paragraph, paragraph.getRuns()));
                }
                List<AbsDataInput<?>> fullDataInfoList = new LinkedList<>();
                for (RunInfo runInfo : runInfoList) {
                    String text = runInfo.getText();
                    if (runInfo.isKey()) {
                        fullDataInfoList.add(this.handleFormat(text));
                        continue;
                    }
                    fullDataInfoList.add(TextDataInput.to(text));
                }
                AbsDataInput finalFullDataInfo = fullDataInfoList.size() == 1 ? fullDataInfoList.get(0) : BatchDataInput.to(fullDataInfoList);
                finalFullDataInfo.setKey(key);
                this.fullDataMap.set(finalFullDataInfo);
            }
            this.document.removeBodyElement(i);
            this.document.removeBodyElement(i);
            break;
        }
    }

    private void fill() throws FullException {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start("填充数据");
        // 获取填充数据
        String handleType = this.templateInfo.getHandleType();
        this.inputDataConvert.convert(this.fullDataMap, handleType);
        this.handleVariableTable();
        stopWatch.stop();
        log.info(stopWatch.getLastTaskName() + "获取完成，耗时：" + stopWatch.getLastTaskTimeMillis() + "毫秒");
        // 填充图表数据
        stopWatch.start("图表数据");
        Map<String, XWPFChart> chartMap = ChartUtil.findChart(this.document);
        for (Map.Entry<String, XWPFChart> entry : chartMap.entrySet()) {
            XWPFChart xwpfChart = entry.getValue();
            FullConfig fullConfig = ChartUtil.getFullConfig(xwpfChart);
            AbsDataInput<?> fullDataInfo = this.handleFormat(entry.getKey());
            this.fullData(xwpfChart, fullConfig, fullDataInfo);
        }
        stopWatch.stop();
        log.info(stopWatch.getLastTaskName() + "填充完成，耗时：" + stopWatch.getLastTaskTimeMillis() + "毫秒");
        // 填充页眉数据
        XWPFHeaderFooterPolicy policy = document.getHeaderFooterPolicy();
        this.fullHeaderFooter(policy.getDefaultHeader(), stopWatch, "页眉数据");
        // 填充页脚数据
        this.fullHeaderFooter(policy.getDefaultFooter(), stopWatch, "页脚数据");
        // 填充段落数据
        stopWatch.start("段落数据");
        ParagraphCreateFun paragraphCreateFun = new ParagraphCreateFun(this.document::insertNewParagraph, this.document::createParagraph);
        this.handleParagraph(0, this.document.getParagraphs(), paragraphCreateFun, null);
        stopWatch.stop();
        log.info(stopWatch.getLastTaskName() + "填充完成，耗时：" + stopWatch.getLastTaskTimeMillis() + "毫秒");
        // 填充表格数据
        stopWatch.start("表格数据");
        List<XWPFTable> tables = TableUtil.findTableList(this.document);
        for (XWPFTable xwpfTable : tables) {
            TableUtil.clearTableName(xwpfTable);
            // 获取表格描述信息
            String tableDescription = TableUtil.getTableDescription(xwpfTable);
            TableUtil.clearTableDescription(xwpfTable);
            FullConfig fullConfig = FullConfig.to(tableDescription);
            this.handleTableRow(xwpfTable, fullConfig);
        }
        stopWatch.stop();
        log.info(stopWatch.getLastTaskName() + "填充完成，耗时：" + stopWatch.getLastTaskTimeMillis() + "毫秒");
        stopWatch.start("自定义");
        this.inputDataConvert.customizeFull(this.document, this.fullDataMap, handleType);
        stopWatch.stop();
        log.info(stopWatch.getLastTaskName() + "填充完成，耗时：" + stopWatch.getLastTaskTimeMillis() + "毫秒");
        log.info("数据填充完成，总耗时：" + stopWatch.getTotalTimeMillis() + "毫秒");
    }

    public void fullHeaderFooter(XWPFHeaderFooter headerFooter, StopWatch stopWatch, String taskName) throws FullException {
        stopWatch.start(taskName);
        if (headerFooter != null) {
            ParagraphCreateFun footerParagraphCreateFun = new ParagraphCreateFun(headerFooter::insertNewParagraph, headerFooter::createParagraph);
            this.handleParagraph(0, headerFooter.getParagraphs(), footerParagraphCreateFun, null);
        }
        stopWatch.stop();
        log.info(stopWatch.getLastTaskName() + "填充完成，耗时：" + stopWatch.getLastTaskTimeMillis() + "毫秒");
    }

    /**
     * 处理表格行数据
     *
     * @param table      表格
     * @param fullConfig 填充配置信息
     */
    public void handleTableRow(XWPFTable table, FullConfig fullConfig) throws FullException {
        final List<XWPFTableRow> rows = table.getRows();
        int rowsSize = rows.size();
        Map<String, XWPFTableRow> fullKeyRowIndexMap = new HashMap<>(8);
        Map<XWPFTableRow, List<String>> rowFullKeyMap = new HashMap<>(8);
        // 原始表格高度信息
        TableHeightInfo sourceTableHeightInfo = new TableHeightInfo(table);
        for (int index = rowsSize - 1; index >= 0; index--) {
            XWPFTableRow row = rows.get(index);
            List<String> fullKeyList;
            try {
                fullKeyList = this.handleTableCell(table, index, row, fullConfig);
                rowFullKeyMap.put(row, fullKeyList);
                for (String fullKey : fullKeyList) {
                    fullKeyRowIndexMap.put(fullKey, row);
                }
            } catch (DelRowException delRowException) {
                table.removeRow(index);
            }
        }
        Set<FullConfig.Rate> autoHeightKeys = fullConfig.getAutoHeightKeys();
        if (autoHeightKeys.isEmpty()) {
            return;
        }
        float tableHeight = fullConfig.getTableHeight();
        float oldTableHeight = sourceTableHeightInfo.getTableSetHeight();
        sourceTableHeightInfo.update(table);
        sourceTableHeightInfo.initFullKeyRowHeightInfo(rowFullKeyMap);
        float newTableHeight = sourceTableHeightInfo.getTableMaxHeight();
        float diff = (tableHeight > 0 ? TableUtil.cm2dxa(tableHeight) : oldTableHeight) - newTableHeight;
        if (diff == 0) {
            return;
        }
        Iterator<FullConfig.Rate> autoHeightKeysIterator = autoHeightKeys.iterator();
        float rateTotal = 0, zeroRateCount = 0;
        while (autoHeightKeysIterator.hasNext()) {
            FullConfig.Rate rate = autoHeightKeysIterator.next();
            String rateName = rate.getName();
            float rateRate = rate.getRate();
            RowHeightInfo rowHeightInfo = sourceTableHeightInfo.getRowHeightInfo(rateName);
            boolean isMoreThanSetHeight = diff < 0 && rowHeightInfo != null && rowHeightInfo.isMoreThanSetHeight();
            if (StrUtil.isEmpty(rateName) || !fullKeyRowIndexMap.containsKey(rateName) || rateRate < 0 || isMoreThanSetHeight) {
                autoHeightKeysIterator.remove();
                continue;
            }
            if (rateRate == 0) {
                zeroRateCount++;
            }
            rateTotal += rateRate;
        }
        float rateDiff = 1F - rateTotal;
        float rateAve = rateDiff / zeroRateCount;
        for (FullConfig.Rate autoHeightKey : autoHeightKeys) {
            XWPFTableRow row = fullKeyRowIndexMap.get(autoHeightKey.getName());
            if (row == null) {
                continue;
            }
            float rate = autoHeightKey.getRate();
            RowHeightInfo rowHeightInfo = sourceTableHeightInfo.getRowHeightInfo(row);
            float maxHeight = rowHeightInfo.getMaxHeight();
            float height = maxHeight + (diff * (rate > 0F ? rate : rateAve));
            TableUtil.setCTHeight(row, height <= 0 ? 0 : ((int) height));
        }
    }

    private int forFull(XWPFTable table, int rowIndex, int cellIndex, String text, FullConfig fullConfig) {
        XWPFTableRow row = table.getRow(rowIndex);
        List<XWPFTableCell> tableCells = row.getTableCells();
        int tableCellsSize = tableCells.size();
        ForDataInput.ForInfo forInfo = ForDataInput.parseForInfo(text);
        Map<Integer, String> titleIndexMap = new LinkedHashMap<>();
        int lastCellIndex = cellIndex;
        for (; lastCellIndex < tableCellsSize; lastCellIndex++) {
            String tText = tableCells.get(lastCellIndex).getText();
            Matcher matcher = ForDataInput.FIELD_PATTERN.matcher(tText);
            if (!matcher.find()) {
                continue;
            }
            String title = matcher.group(1);
            titleIndexMap.put(lastCellIndex, title);
            if (tText.contains("}")) {
                break;
            }
        }
        List<ForDataInput> fullDataInfoList = this.fullDataMap.get(forInfo.getKey(), ForDataInput.class);
        FullContent fullContent = new FullContent(rowIndex, cellIndex);
        fullContent.setTable(table);
        for (ForDataInput fullDataInfo : fullDataInfoList) {
            fullDataInfo.setFullConfig(fullConfig);
            fullDataInfo.setTitleIndexMap(titleIndexMap);
            fullDataInfo.setForInfo(forInfo);
            fullDataInfo.full(fullContent);
        }
        return lastCellIndex;
    }

    /**
     * 处理表格单元格数据
     *
     * @param row        表格行数据
     * @param fullConfig 填充配置信息
     * @return 填充字段
     * @throws RuntimeException 异常
     */
    public List<String> handleTableCell(XWPFTable table, int rowIndex, XWPFTableRow row, FullConfig fullConfig) throws RuntimeException, FullException {
        final List<XWPFTableCell> tableCells = row.getTableCells();
        List<String> fullKeyList = new LinkedList<>();
        int tableCellsSize = tableCells.size();
        for (int i = 0; i < tableCellsSize; i++) {
            XWPFTableCell tableCell = tableCells.get(i);
            int width = -1;
            try {
                width = tableCell.getWidth();
            } catch (NullPointerException ignored) { }
            String text = tableCell.getText();
            if (ForDataInput.check(text)) {
                int lastCellIndex = this.forFull(table, rowIndex, i, text, fullConfig);
                if (lastCellIndex < tableCellsSize) {
                    i = lastCellIndex;
                    continue;
                }
                break;
            }
            ParagraphCreateFun paragraphCreateFun = new ParagraphCreateFun(tableCell::insertNewParagraph, tableCell::addParagraph);
            List<String> tempList = this.handleParagraph(width, tableCell.getParagraphs(), paragraphCreateFun, fullConfig);
            fullKeyList.addAll(tempList);
        }
        return fullKeyList;
    }

    /**
     * 处理段落信息
     *
     * @param width              段落所在区域的宽度
     * @param paragraphs         段落列表
     * @param paragraphCreateFun 段落插入方法
     * @param fullConfig         填充配置信息
     * @return 填充字段
     * @throws RuntimeException 异常
     */
    public List<String> handleParagraph(float width, List<XWPFParagraph> paragraphs, ParagraphCreateFun paragraphCreateFun, FullConfig fullConfig) throws RuntimeException, FullException {
        if (paragraphs.isEmpty()) {
            return Collections.emptyList();
        }
        int paragraphsSize = paragraphs.size();
        List<String> fullKeyList = new LinkedList<>();
        for (int index = paragraphsSize - 1; index >= 0; index--) {
            XWPFParagraph xwpfParagraph = paragraphs.get(index);
            XmlCursor xmlCursor;
            if (index < paragraphsSize - 1) {
                xmlCursor = paragraphs.get(index + 1).getCTP().newCursor();
            } else {
                xmlCursor = null;
            }
            paragraphCreateFun.setXmlCursor(xmlCursor);
            List<String> tempList = this.handleParagraph(width, xwpfParagraph, paragraphCreateFun, fullConfig);
            fullKeyList.addAll(tempList);
        }
        return fullKeyList;
    }

    /**
     * 处理段落信息
     *
     * @param width              段落所在区域的宽度
     * @param paragraph          段落
     * @param paragraphCreateFun 段落添加方法
     * @param fullConfig         填充配置信息
     * @return 填充字段
     */
    public List<String> handleParagraph(float width, XWPFParagraph paragraph, ParagraphCreateFun paragraphCreateFun, FullConfig fullConfig) throws FullException {
        List<XWPFRun> runs = paragraph.getRuns();
        List<RunInfo> runInfoList = getRunInfo(paragraph, runs);
        if (runInfoList.isEmpty()) {
            return Collections.emptyList();
        }
        int runsSize = runs.size();
        for (int i = 0; i < runsSize; i++) {
            paragraph.removeRun(0);
        }
        List<String> fullKeyList = new LinkedList<>();
        XWPFParagraph curParagraph = paragraph;
        for (RunInfo runInfo : runInfoList) {
            String text = runInfo.getText();
            AbsDataInput<?> fullDataInfo;
            if (runInfo.isKey()) {
                fullDataInfo = this.handleFormat(text);
                fullKeyList.add(fullDataInfo.getKey());
            } else {
                fullDataInfo = TextDataInput.to(text);
            }
            fullDataInfo.setFullConfig(Optional.ofNullable(fullConfig).orElse(FullConfig.empty()));
            fullDataInfo.setFullDataMap(this.fullDataMap);
            CTRPr ctrPr = runInfo.getCtrPr();
            int fontSize = 24;
            if (ctrPr != null) {
                CTHpsMeasure[] szCsArray = ctrPr.getSzCsArray();
                if (szCsArray != null && szCsArray.length > 0) {
                    Object val = szCsArray[0].getVal();
                    fontSize = ((BigInteger) val).intValue();
                }
            }
            FullContent fullContent = new FullContent(-1, -1);
            fullContent.setParagraph(curParagraph);
            fullContent.setParagraphCreateFun(paragraphCreateFun);
            fullContent.setCtrPr(ctrPr);
            fullContent.setWidth(width);
            fullContent.setFontSize(fontSize);
            if (fullDataInfo instanceof BatchDataInput) {
                BatchDataInput batchDataInput = (BatchDataInput) fullDataInfo;
                batchDataInput.full(fullContent);
                curParagraph = fullContent.getParagraph();
                continue;
            }
            boolean setIndentationFirstLine = false;
            if (runInfo.isKey()) {
                setIndentationFirstLine = fullDataInfo.cacheAutoIndentationFirstLine(width, fontSize);
            }
            curParagraph = TableUtil.full(fullDataInfo, fullContent, setIndentationFirstLine);
        }
        return fullKeyList;
    }

    public List<RunInfo> getRunInfo(XWPFParagraph paragraph, List<XWPFRun> runs) {
        if (runs == null || runs.isEmpty()) {
            return Collections.emptyList();
        }
        // 判断是否需要填充
        String runsStr = runs.stream().map(XWPFRun::text).collect(Collectors.joining());
        if (!runsStr.contains("${") && !runsStr.contains("}")) {
            return Collections.emptyList();
        }
        int runsSize = runs.size();
        List<RunInfo> runInfoList = new LinkedList<>();
        RunInfo lastRunInfo = null;
        boolean format = false;
        StringBuilder formatKeyStringBuilder = new StringBuilder();
        for (int i = 0; i < runsSize; i++) {
            XWPFRun run = runs.get(i);
            String runText = run.text();
            int textLength = runText.length();
            CTRPr ctrPr = TableUtil.copyRunStyle(paragraph, i);
            for (int textIndex = 0; textIndex < textLength; textIndex++) {
                char curChar = runText.charAt(textIndex);
                if (!format && curChar == '$') {
                    String text = formatKeyStringBuilder.toString();
                    if (StrUtil.isNotEmpty(text)) {
                        runInfoList.add(RunInfo.to(text, ctrPr));
                    }
                    lastRunInfo = new RunInfo(ctrPr, i, true);
                    runInfoList.add(lastRunInfo);
                    formatKeyStringBuilder = new StringBuilder();
                    format = true;
                }
                formatKeyStringBuilder.append(curChar);
                if (format && curChar == '}') {
                    lastRunInfo.setText(formatKeyStringBuilder.toString());
                    formatKeyStringBuilder = new StringBuilder();
                    format = false;
                }
            }
            String text = formatKeyStringBuilder.toString();
            if (!format && StrUtil.isNotEmpty(text)) {
                runInfoList.add(RunInfo.to(text, ctrPr));
                formatKeyStringBuilder = new StringBuilder();
            }
        }
        return runInfoList;
    }

    /**
     * 解析从模板中获取到键的格式
     *
     * @param formatKey 键
     * @return 填充数据列表
     */
    public AbsDataInput<?> handleFormat(String formatKey) throws FullException {
        AbsDataInput<?> absDataInput = ExtendDataInputUtil.find(formatKey);
        if (absDataInput != null) {
            return absDataInput;
        }
        int length = formatKey.length();
        String sourceKey = formatKey.substring(2, length - 1);
        String dataFormat = "";
        if (sourceKey.contains(":")) {
            String[] split = sourceKey.split(":");
            sourceKey = split[0];
            dataFormat = split[1];
        }
        AbsDataInput<?> fullDataInfo = null;
        if (sourceKey.contains("|")) {
            String[] keys = sourceKey.split("\\|");
            for (String key : keys) {
                fullDataInfo = this.fullDataMap.get(key);
                if (ObjUtil.isNotEmpty(fullDataInfo) && fullDataInfo.isNotEmpty()) {
                    break;
                }
            }
        } else {
            fullDataInfo = this.fullDataMap.get(sourceKey);
        }
        if (fullDataInfo == null) {
            fullDataInfo = TextDataInput.to("").setKey(sourceKey);
        }
        if (fullDataInfo instanceof TextDataInput) {
            ((TextDataInput) fullDataInfo).setDataFormat(dataFormat);
        }
        return fullDataInfo;
    }

    /**
     * 保存文件
     *
     * @return 结果文件
     */
    public File save() throws FullException {
        try (OutputStream outputStream = Files.newOutputStream(Paths.get(targetFileFullPath))) {
            this.document.write(outputStream);
            outputStream.close();
            return FileUtil.file(targetFileFullPath);
        } catch (IOException e) {
            throw new FullException("文件保存失败。", e);
        }
    }

    @Override
    public void close() {
        this.inputDataConvert.close();
        IoUtil.close(this.document);
    }

    private void fullData(Object obj, FullConfig fullConfig, AbsDataInput<?> fullDataInfo) throws FullException {
        if (fullDataInfo instanceof BatchDataInput) {
            BatchDataInput batchDataInput = (BatchDataInput) fullDataInfo;
            List<AbsDataInput<?>> list = batchDataInput.getList();
            for (AbsDataInput<?> item : list) {
                this.fullData(obj, item, fullConfig);
            }
            return;
        }
        this.fullData(obj, fullDataInfo, fullConfig);
    }

    private void fullData(Object obj, AbsDataInput<?> fullDataInfo, FullConfig fullConfig) throws FullException {
        if (obj instanceof XWPFTable) {
            TableUtil.fullData(((XWPFTable) obj), fullConfig, fullDataInfo);
        } else if (obj instanceof XWPFChart) {
            ChartUtil.fullData(((XWPFChart) obj), fullConfig, fullDataInfo);
        }
    }

    public String getTargetFilePath() {
        return targetFilePath;
    }

    public String getTargetFileName() {
        return targetFileName;
    }

    public String getTargetFileFullPath() {
        return targetFileFullPath;
    }
}
