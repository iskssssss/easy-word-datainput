package top.kongsheng.common.word.datainput.utils;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.word.datainput.core.FullDataMap;
import top.kongsheng.common.word.datainput.model.input.AbsDataInput;
import top.kongsheng.common.word.datainput.model.input.BatchDataInput;
import top.kongsheng.common.word.datainput.model.input.TextDataInput;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

/**
 * 填充数据工具类
 *
 * @author 孔胜
 * @date 2023/8/3 11:06
 */
public class FullDataInfoUtil {

    public static List<AbsDataInput<?>> splitValue(TextDataInput textDataInput) {
        if (textDataInput == null) {
            return Collections.emptyList();
        }
        String valueStr = textDataInput.getValueStr();
        if (StrUtil.isEmpty(valueStr)) {
            ArrayList<AbsDataInput<?>> absDataInputs = new ArrayList<>();
            absDataInputs.add(textDataInput);
            return absDataInputs;
        }
        String key = textDataInput.getKey();
        boolean newLine = textDataInput.isNewLine();
        String replace = valueStr.replace("<br />", "\n").replace("<br/>", "\n").replace("\\n", "\n");
        if (replace.contains("\n")) {
            List<AbsDataInput<?>> result = new LinkedList<>();
            String[] valueSplit = replace.split("\n");
            int valueSplitLength = valueSplit.length;
            for (int i = 0; i < valueSplitLength; i++) {
                String item = valueSplit[i];
                TextDataInput fullDataInfo = textDataInput.copy(key, item);
                result.add(fullDataInfo.setNewLine(i > 0 || newLine)
                        .setIndentationFirstLine(textDataInput.getIndentationFirstLine())
                        .setAutoIndentationFirstLine(textDataInput.isAutoIndentationFirstLine())
                );
            }
            return result;
        }
        ArrayList<AbsDataInput<?>> absDataInputs = new ArrayList<>();
        absDataInputs.add(textDataInput);
        return absDataInputs;
    }

    public static void splitValue(FullDataMap fullDataMap, String key, int indentationFirstLine, boolean autoIndentationFirstLine) {
        AbsDataInput<?> sourceFullDataInfo = fullDataMap.get(key);
        if (sourceFullDataInfo == null) {
            return;
        }
        sourceFullDataInfo.setIndentationFirstLine(indentationFirstLine);
        sourceFullDataInfo.setAutoIndentationFirstLine(autoIndentationFirstLine);
        if (sourceFullDataInfo instanceof BatchDataInput) {
            BatchDataInput batchDataInput = (BatchDataInput) sourceFullDataInfo;
            batchDataInput.split();
            return;
        }
        if (!(sourceFullDataInfo instanceof TextDataInput) || sourceFullDataInfo.isEmpty()) {
            return;
        }
        List<AbsDataInput<?>> textFullDataInfoList = sourceFullDataInfo.split();
        if (textFullDataInfoList.isEmpty()) {
            return;
        }
        BatchDataInput batchDataInput = new BatchDataInput();
        batchDataInput.setKey(key);
        batchDataInput.addAll(textFullDataInfoList);
        fullDataMap.put(key, batchDataInput);
    }
}
