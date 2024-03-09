## easy-poi

基于预设模板的word导出工具，实现了word自动生成和导出功能。提供了可配置的导出模板，满足了不同业务场景下的数据导出需求。

## 示例

```java
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
```

模板

<img src=".\images\template-image.png" alt="template-image.png" />

填充结果

<img src=".\images\result-image.png" alt="template-image.png" />

