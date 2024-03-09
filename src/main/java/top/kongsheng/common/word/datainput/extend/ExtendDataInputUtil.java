package top.kongsheng.common.word.datainput.extend;

import top.kongsheng.common.word.datainput.model.input.AbsDataInput;
import lombok.extern.log4j.Log4j2;

import java.util.ArrayList;
import java.util.List;
import java.util.ServiceLoader;

/**
 * 用于扩展插件工具类
 *
 * @author 孔胜
 * @date 2024/1/11 15:06
 */
@Log4j2
public class ExtendDataInputUtil {
    private static final List<ExtendDataInput> EXTEND_FULL_DATA_LIST = new ArrayList<>();

    static {
        ServiceLoader<ExtendDataInput> extendFullDataServiceLoader = ServiceLoader.load(ExtendDataInput.class);
        for (ExtendDataInput extendDataInput : extendFullDataServiceLoader) {
            log.info("loading data input extend:" + extendDataInput.getClass());
            EXTEND_FULL_DATA_LIST.add(extendDataInput);
        }
    }

    public static AbsDataInput<?> find(String key) {
        for (ExtendDataInput extendDataInput : EXTEND_FULL_DATA_LIST) {
            if (extendDataInput.check(key)) {
                return extendDataInput.to(key);
            }
        }
        return null;
    }
}
