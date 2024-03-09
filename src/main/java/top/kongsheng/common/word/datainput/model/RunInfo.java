package top.kongsheng.common.word.datainput.model;

import cn.hutool.core.util.StrUtil;
import lombok.Data;
import lombok.ToString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

/**
 * 文本区域 数据
 *
 * @author 孔胜
 * @date 2023/8/2 17:27
 */
@Data
@ToString
public class RunInfo {

    /**
     * 文本
     */
    private String text;
    /**
     * 文本样式
     */
    private CTRPr ctrPr;
    /**
     * 开始位置
     */
    private int position;
    /**
     * 是否是填充键
     */
    private boolean key = false;

    public RunInfo() {
    }

    public RunInfo(String text, boolean key) {
        this.text = text;
        this.key = key;
    }

    public RunInfo(CTRPr ctrPr, int position, boolean key) {
        this.ctrPr = ctrPr;
        this.position = position;
        this.key = key;
    }

    public String getFullFormatKey() {
        return StrUtil.format("${{}}", this.text);
    }

    public static RunInfo to(String text, CTRPr ctrPr) {
        RunInfo runInfo = new RunInfo();
        runInfo.setText(text);
        runInfo.setCtrPr(ctrPr);
        return runInfo;
    }

}
