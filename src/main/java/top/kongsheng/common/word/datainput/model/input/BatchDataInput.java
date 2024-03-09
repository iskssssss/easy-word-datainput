package top.kongsheng.common.word.datainput.model.input;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.word.datainput.exception.DelRowException;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.config.FullConfig;
import top.kongsheng.common.word.datainput.utils.TableUtil;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.LinkedList;
import java.util.List;
import java.util.Optional;
import java.util.Set;

/**
 * 批量填充数据
 *
 * @author 孔胜
 * @date 2023/9/1 14:46
 */
public class BatchDataInput extends AbsDataInput<BatchDataInput> {

    private final List<AbsDataInput<?>> list = new LinkedList<>();

    public BatchDataInput() {
        super(-1);
    }

    public List<AbsDataInput<?>> getList() {
        return list;
    }

    @Override
    public BatchDataInput setKey(String key) {
        super.setKey(key);
        list.forEach(item -> {
            if (StrUtil.isEmpty(item.getKey())) {
                item.setKey(key);
            }
        });
        return this;
    }

    public BatchDataInput addAll(List<AbsDataInput<?>> fullDataInfoList) {
        for (AbsDataInput<?> absDataInput : fullDataInfoList) {
            this.add(absDataInput);
        }
        return this;
    }

    public BatchDataInput add(AbsDataInput<?> fullDataInfo) {
        return this.add(-1, fullDataInfo);
    }

    public BatchDataInput add(int i, AbsDataInput<?> fullDataInfo) {
        if (StrUtil.isEmpty(fullDataInfo.getKey())) {
            fullDataInfo.setKey(getKey());
        }
        if (i < 0) {
            list.add(fullDataInfo);
            return this;
        }
        list.add(i, fullDataInfo);
        return this;
    }

    @Override
    public long dataLength() {
        return list.size();
    }

    @Override
    public boolean isEmpty() {
        for (AbsDataInput<?> fullDataInfo : list) {
            if (fullDataInfo.isNotEmpty()) {
                return false;
            }
        }
        return true;
    }

    @Override
    public boolean cacheAutoIndentationFirstLine(float width, float fontSize) {
        boolean setIndentationFirstLine = false;
        for (AbsDataInput<?> fullDataInfo : list) {
            boolean cacheAutoIndentationFirstLine = fullDataInfo.cacheAutoIndentationFirstLine(width, fontSize);
            setIndentationFirstLine = setIndentationFirstLine || cacheAutoIndentationFirstLine;
        }
        return setIndentationFirstLine;
    }

    @Override
    public void full(FullContent fullContent) throws FullException {
        XWPFParagraph curParagraph = fullContent.getParagraph();
        float width = fullContent.getWidth();
        float fontSize = fullContent.getFontSize();
        String key = super.getKey();
        FullConfig fullConfig = getFullConfig();
        FullConfig finalFullConfig = Optional.ofNullable(fullConfig).orElse(FullConfig.empty());
        Set<String> emptyDelRowKey = finalFullConfig.getEmptyDelRowKey();
        boolean containsKey = emptyDelRowKey.contains(key);
        if (this.isEmpty() && containsKey) {
            throw new DelRowException();
        }
        boolean emptyDelRowKeyEmpty = emptyDelRowKey.isEmpty();
        emptyDelRowKey.remove(key);
        boolean setIndentationFirstLine = this.cacheAutoIndentationFirstLine(width, fontSize);
        for (AbsDataInput<?> fullDataInfo : list) {
            fullDataInfo.setFullConfig(finalFullConfig);
            fullDataInfo.setFullDataMap(super.getFullDataMap());
            curParagraph = TableUtil.full(fullDataInfo, fullContent, setIndentationFirstLine);
            fullContent.setParagraph(curParagraph);
        }
        if (emptyDelRowKeyEmpty || !containsKey) {
            return;
        }
        emptyDelRowKey.add(key);
    }


    @Override
    public List<AbsDataInput<?>> split() {
        List<AbsDataInput<?>> t = new LinkedList<>();
        for (AbsDataInput<?> fullDataInfo : list) {
            fullDataInfo.setIndentationFirstLine(this.getIndentationFirstLine());
            fullDataInfo.setAutoIndentationFirstLine(this.isAutoIndentationFirstLine());
            t.addAll(fullDataInfo.split());
        }
        list.clear();
        list.addAll(t);
        return t;
    }

    public static BatchDataInput to(List<AbsDataInput<?>> fullDataInfoList) {
        BatchDataInput batchDataInput = new BatchDataInput();
        if (fullDataInfoList == null || fullDataInfoList.isEmpty()) {
            return batchDataInput;
        }
        batchDataInput.addAll(fullDataInfoList);
        return batchDataInput;
    }

    public static BatchDataInput to(AbsDataInput... fullDataInfoList) {
        BatchDataInput batchDataInput = new BatchDataInput();
        if (fullDataInfoList == null || fullDataInfoList.length < 1) {
            return batchDataInput;
        }
        for (AbsDataInput t : fullDataInfoList) {
            if (t == null) {
                continue;
            }
            batchDataInput.add(t);
        }
        return batchDataInput;
    }
}
