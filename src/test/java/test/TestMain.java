package test;

import cn.hutool.core.io.IoUtil;
import cn.hutool.core.util.RandomUtil;
import lombok.Data;
import lombok.ToString;
import org.junit.Test;
import top.kongsheng.common.word.datainput.WordDataInputManage;
import top.kongsheng.common.word.datainput.convert.InputDataConvert;
import top.kongsheng.common.word.datainput.core.FullDataMap;
import top.kongsheng.common.word.datainput.core.WordDataInputCore;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.TemplateInfo;
import top.kongsheng.common.word.datainput.model.input.ChartDataInput;
import top.kongsheng.common.word.datainput.model.input.ForDataInput;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

/**
 * TestMain
 *
 * @author 孔胜
 * @date 2023/8/14 17:41
 */
public class TestMain {

    @Data
    @ToString
    static class TestModel {
        private Integer index;

        private String name;
    }

    @Test
    public void test() throws IOException {
        WordDataInputManage.setWordFataInputFileSaveDir("D:\\exports");
        InputDataConvert inputDataConvert = new InputDataConvert() {
            @Override
            public void convert(FullDataMap fullDataMap, String handleType) throws FullException {
                fullDataMap.set("段落填充", "段落填充数据");
                fullDataMap.set("段落样式保留", "段落样式保留数据");
                fullDataMap.set("表格填充1", "表格填充1数据");
                fullDataMap.set("表格填充2", "表格填充2数据");
                // 列表填充
                ForDataInput forDataInput = new ForDataInput("list");
                List<TestModel> testModelList = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    TestModel testModel = new TestModel();
                    testModel.setIndex(i + 1);
                    testModel.setName("张" + testModel.getIndex());
                    testModelList.add(testModel);
                }
                forDataInput.setDataList(testModelList);
                fullDataMap.set(forDataInput);
                // 表单填充
                ChartDataInput chartDataInput = new ChartDataInput("chart");
                String[] titles = new String[]{"商品1", "商品2", "商品3"};
                for (int i = 1; i <= 12; i++) {
                    chartDataInput.addCategories(i + "月");
                    for (String title : titles) {
                        double val = RandomUtil.randomDouble(20, 100);
                        chartDataInput.addTitleValue(title, val);
                    }
                }
                fullDataMap.set(chartDataInput);
            }
        };
        TemplateInfo templateInfo = new TemplateInfo();
        URL resource = TestMain.class.getResource("/template.docx");
        if (resource == null) {
            return;
        }
        try(InputStream inputStream = resource.openStream()) {
            templateInfo.setFileName("template.docx");
            templateInfo.setFileBytes(IoUtil.readBytes(inputStream));
            try (WordDataInputCore wordDataInputCore = new WordDataInputCore(templateInfo)) {
                wordDataInputCore.setInputDataConvert(inputDataConvert);
                wordDataInputCore.run();
                File printFile = wordDataInputCore.save();
                if (printFile == null) {
                    throw new SecurityException("，文件保存失败。");
                }
            } catch (Exception e) {
                throw new RuntimeException("数据填充失败。", e);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}

