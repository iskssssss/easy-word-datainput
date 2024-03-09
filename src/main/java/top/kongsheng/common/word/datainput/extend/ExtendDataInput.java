package top.kongsheng.common.word.datainput.extend;

import top.kongsheng.common.word.datainput.model.input.AbsDataInput;

/**
 * 用于扩展插件
 *
 * @author 孔胜
 * @date 2024/1/11 15:06
 */
public interface ExtendDataInput {

    /**
     * 校验是否符合
     *
     * @param key 键
     * @return 是否符合
     */
    boolean check(String key);

    /**
     * 返回填充类
     *
     * @param key 键
     * @return 填充类
     */
    AbsDataInput<?> to(String key);
}
