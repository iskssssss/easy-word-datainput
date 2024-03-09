package test.model;

import top.kongsheng.common.word.datainput.model.input.ChartDataInput;
import top.kongsheng.common.word.datainput.model.input.ForDataInput;
import lombok.Data;

import java.io.Serializable;
import java.util.*;
import java.util.stream.Collectors;

@Data
public class TaskInfoLeadershipBaseVo implements Serializable {

    private String id;

    private String time;

    private String reportNumber;

    private String summarize;

    private String fileName;

    private TaskInfoLeadershipCountVo leadershipCountVo;

    private List<TaskInfoLeadershipSourceByEchart> sourceByEchart;

    private LinkedHashMap<String, List<IntegerKeyValueWithId>> sourceByTypes = new LinkedHashMap<>();

    private List<TaskInfoInstructionObject> infoInstructionObjects;


    private List<TaskInfoLeadershipSourceByEchart> instructTypeByEchart;

    private List<TaskInfoTagsVo> secondaryInstruction;

    private List<TaskInfoTagsVo> firstInstructionInstruction;

    private List<TaskInfoTagsVo> supervision;

    private List<TaskInfoTagsVo> standingCommittee;

    private List<TaskInfoTagsVo> office;

    private String analysis;

    private String createBy;

    private List<TaskInfoLeadershipRate> undertakeTaskDept;

    private static final int[][] QUARTERS = new int[][]{
            {},
            {1, 2, 3},
            {4, 5, 6},
            {7, 8, 9},
            {10, 11, 12}
    };

    private int[] getRange(int type, int num, int sourceSize) {
        int startIndex = 0, endIndex = sourceSize;
        if (sourceSize < 12) {
            return new int[]{startIndex, endIndex};
        }
        switch (type) {
            case 3:
                startIndex = num - 1;
                endIndex = num;
                break;
            case 2:
                startIndex = QUARTERS[num][0] - 1;
                endIndex = QUARTERS[num][2];
                break;
            case 1:
            default:
                break;
        }
        return new int[]{startIndex, endIndex};
    }

    public ChartDataInput createSourceCharData(int type, int num) {
        ChartDataInput chartDataInput = new ChartDataInput("来文单位或个人");
        int sourceSize = sourceByEchart.size();
        int[] ranges = getRange(type, num, sourceSize);
        int startIndex = ranges[0], endIndex = ranges[1];
        Set<String> keySet = null;
        for (int i = startIndex; i < endIndex; i++) {
            TaskInfoLeadershipSourceByEchart leadershipSource = sourceByEchart.get(i);
            Map<String, Integer> data = leadershipSource.getData();
            chartDataInput.addCategories(leadershipSource.getMonth());
            keySet = data.keySet();
            for (String title : keySet) {
                chartDataInput.addTitleValue(title, Double.valueOf(data.get(title)));
            }
        }
        if (keySet == null) {
            return chartDataInput;
        }
        int range = (sourceSize - (endIndex - startIndex)) / 2;
        if (range == 0) {
            return chartDataInput;
        }
        for (int i = 0; i < range; i++) {
            chartDataInput.insertCategories(0, "");
            chartDataInput.addCategories("");
            for (String title : keySet) {
                chartDataInput.insertTitleValue(title, 0, (double) 0);
                chartDataInput.addTitleValue(title, (double) 0);
            }
        }
        return chartDataInput;
    }

    public List<ForDataInput> createSourceTableData() {
        final String formatKey = "来文单位或个人";
        List<ForDataInput> fullDataInfos = new LinkedList<>();
        int maxRowTotal = 0;
        for (Map.Entry<String, List<IntegerKeyValueWithId>> entry : sourceByTypes.entrySet()) {
            List<IntegerKeyValueWithId> valueList = entry.getValue();
            ForDataInput forDataInput = new ForDataInput(entry.getKey());
            forDataInput.setExtendKey(formatKey);
            forDataInput.setDataList(valueList);
            maxRowTotal = Math.max(maxRowTotal, valueList.size());
            fullDataInfos.add(forDataInput);
        }
        for (ForDataInput forDataInput : fullDataInfos) {
            forDataInput.setMaxRowNum(maxRowTotal);
        }
        return fullDataInfos;
    }

    public List<ChartDataInput> createInfoInstructionObjectsCharData() {
        List<ChartDataInput> chartDataInputList = new LinkedList<>();
        for (TaskInfoInstructionObject infoInstructionObject : infoInstructionObjects) {
            String name = infoInstructionObject.getName();
            ChartDataInput chartDataInput = new ChartDataInput(name);
            chartDataInputList.add(chartDataInput);
            List<TaskInfoLeadershipRate> rateList = infoInstructionObject.getData();
            for (TaskInfoLeadershipRate datum : rateList) {
                chartDataInput.addCategories(datum.getKey());
                chartDataInput.addTitleValue(name, Double.valueOf(datum.getValue()));
            }
        }
        return chartDataInputList;
    }

    public ChartDataInput createInstructTypeCharData(int type, int num) {
        ChartDataInput chartDataInput = new ChartDataInput("批示类型");
        int sourceSize = instructTypeByEchart.size();
        int[] range = getRange(type, num, sourceSize);
        int startIndex = range[0], endIndex = range[1];
        for (int i = startIndex; i < endIndex; i++) {
            TaskInfoLeadershipSourceByEchart leadershipSource = instructTypeByEchart.get(i);
            Map<String, Integer> data = leadershipSource.getData();
            chartDataInput.addCategories(leadershipSource.getMonth());
            Set<String> keySet = data.keySet();
            for (String title : keySet) {
                Integer value = data.get(title);
                chartDataInput.addTitleValue(title, Double.valueOf(value));
            }
        }
        return chartDataInput;
    }

    public List<ForDataInput> createSecondaryInstructionTableData() {
        String key = "批示类型";
        List<ForDataInput> tableFullDataInfoList = new LinkedList<>();
        ForDataInput firstInstructionInstructionTableFullDataInfo = new ForDataInput("firstInstructionInstruction");
        firstInstructionInstructionTableFullDataInfo.setExtendKey(key);
        firstInstructionInstructionTableFullDataInfo.setDataList(firstInstructionInstruction);
        tableFullDataInfoList.add(firstInstructionInstructionTableFullDataInfo);
        ForDataInput secondaryInstructionTableFullDataInfo = new ForDataInput("secondaryInstruction");
        secondaryInstructionTableFullDataInfo.setExtendKey(key);
        secondaryInstructionTableFullDataInfo.setDataList(secondaryInstruction);
        tableFullDataInfoList.add(secondaryInstructionTableFullDataInfo);
        return tableFullDataInfoList;
    }

    public ChartDataInput createSupervisionCharData() {
        ChartDataInput chartDataInput = new ChartDataInput("按人大工作分类");
        for (TaskInfoTagsVo taskInfoTagsVo : supervision) {
            String name = taskInfoTagsVo.getName();
            chartDataInput.addCategories(name);
            Integer summarize = taskInfoTagsVo.getSummarize();
            chartDataInput.addTitleValue(chartDataInput.getKey(), Double.valueOf(summarize));
        }
        return chartDataInput;
    }

    public ForDataInput createSupervisionTableData() {
        ForDataInput forDataInput = new ForDataInput("supervision");
        List<TaskInfoTagsVo> infoTagsVoList = supervision.stream().sorted((o1, o2) -> {
            int summarize1 = Optional.ofNullable(o1.getSummarize()).orElse(0);
            int summarize2 = Optional.ofNullable(o2.getSummarize()).orElse(0);
            return Integer.compare(summarize2, summarize1);
        }).collect(Collectors.toList());
        forDataInput.setDataList(infoTagsVoList);
        return forDataInput;
    }

    public ChartDataInput createOfficeCharData() {
        ChartDataInput chartDataInput = new ChartDataInput("重点热词");
        for (TaskInfoTagsVo taskInfoTagsVo : office) {
            String name = taskInfoTagsVo.getName();
            chartDataInput.addCategories(name);
            Integer summarize = taskInfoTagsVo.getSummarize();
            chartDataInput.addTitleValue(chartDataInput.getKey(), Double.valueOf(summarize));
        }
        return chartDataInput;
    }

    public ForDataInput createOfficeTableData() {
        ForDataInput forDataInput = new ForDataInput("office");
        List<TaskInfoTagsVo> infoTagsVoList = office.stream().filter(item -> {
            Integer summarize = item.getSummarize();
            return summarize != null && summarize > 0;
        }).sorted((o1, o2) -> {
            int summarize1 = Optional.ofNullable(o1.getSummarize()).orElse(0);
            int summarize2 = Optional.ofNullable(o2.getSummarize()).orElse(0);
            return Integer.compare(summarize2, summarize1);
        }).collect(Collectors.toList());
        forDataInput.setDataList(infoTagsVoList);
        return forDataInput;
    }

    public ForDataInput createUndertakeTaskDeptTableData() {
        ForDataInput tableFullDataInfo = new ForDataInput("utdl");
        tableFullDataInfo.setDataList(undertakeTaskDept);
        return tableFullDataInfo;
    }
}
