package top.kongsheng.common.word.datainput.config;

import cn.hutool.json.JSONUtil;
import lombok.Data;
import lombok.ToString;

import java.io.Serializable;
import java.util.*;

/**
 * 填充配置信息
 *
 * @author 孔胜
 * @date 2023/8/14 17:15
 */
@Data
@ToString
public class FullConfig {

    /**
     * 绘制多少数据
     */
    int showDataNum = -1;

    /**
     * 设置是否显示数据标签
     */
    private boolean showDataLabel = true;

    /**
     * 设置是否显示数据点的值（value）。
     * <p>0：不显示 1.显示全部 2.显示大于0的</p>
     */
    private int showType = 2;

    /**
     * 为空值删除所在行
     */
    private Set<String> emptyDelRowKey;

    /**
     * 自动高度
     */
    private Set<Rate> autoHeightKeys;

    /**
     * 自动高度容错值
     */
    private float autoHeightFaultTolerantValue = 0;

    /**
     * 表格高度
     */
    private float tableHeight = 0;

    /**
     * 自动合并坐标列表
     */
    private List<MergeCellInfo> autoMergeCells;

    /**
     * 1.覆盖 2.插入
     */
    private int tableFullType = 1;

    public static FullConfig to(String tableDescription) {
        return Optional.ofNullable(tableDescription.startsWith("{") ? JSONUtil.toBean(tableDescription, FullConfig.class) : null).orElse(FullConfig.empty());
    }

    public Set<String> getEmptyDelRowKey() {
        if (this.emptyDelRowKey == null) {
            this.emptyDelRowKey = Collections.emptySet();
        }
        return this.emptyDelRowKey;
    }

    public Set<Rate> getAutoHeightKeys() {
        if (this.autoHeightKeys == null) {
            this.autoHeightKeys = Collections.emptySet();
        }
        return this.autoHeightKeys;
    }

    public List<MergeCellInfo> getAutoMergeCells() {
        if (this.autoMergeCells == null) {
            this.autoMergeCells = Collections.emptyList();
        }
        return this.autoMergeCells;
    }

    public static FullConfig empty() {
        return new FullConfig();
    }

    @Data
    @ToString
    public static class Rate implements Serializable {
        static final long serialVersionUID = 42L;

        private String name;

        private float rate = 0F;
    }

    @Data
    @ToString
    public static class MergeCellInfo implements Serializable {

        static final long serialVersionUID = 42L;

        private int sx = 0;
        private int sy = 0;

        private int ex = 0;
        private int ey = 0;

        public boolean isEmpty() {
            return sx == 0 && sy == 0 && ex == 0 && ey == 0;
        }
    }
}
