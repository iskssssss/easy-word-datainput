package top.kongsheng.common.word.datainput.model.table;

import lombok.Data;
import lombok.ToString;

import java.io.Serializable;

/**
 * 行高信息
 *
 * @author 孔胜
 * @date 2023/9/13 10:26
 */
@Data
@ToString
public  class RowHeightInfo implements Serializable {
    static final long serialVersionUID = 42L;
    /**
     * 设置高度
     */
    private float setHeight;

    /**
     * 内容高度
     */
    private float contentHeight;

    /**
     * 内容高度是否超过设置高度
     *
     * @return 内容高度是否超过设置高度
     */
    public boolean isMoreThanSetHeight() {
        return contentHeight > setHeight;
    }

    /**
     * 获取最大高度
     *
     * @return 最大高度
     */
    public float getMaxHeight() {
        return Math.max(this.setHeight, this.contentHeight);
    }

    /**
     * 获取最小高度
     *
     * @return 最小高度
     */
    public float getMinHeight() {
        return Math.min(this.setHeight, this.contentHeight);
    }
}
