package top.kongsheng.common.word.datainput.convert;

import top.kongsheng.common.word.datainput.core.FullDataMap;
import top.kongsheng.common.word.datainput.exception.FullException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.Closeable;

/**
 * 填充数据转换接口
 *
 * @author 孔胜
 * @date 2023/7/27 17:09
 */
public abstract class InputDataConvert implements Closeable {

    /**
     * 数据转换
     *
     * @param fullDataMap 参数集合
     * @param handleType  处理类型
     * @return 数据添加
     */
    public abstract void convert(FullDataMap fullDataMap, String handleType) throws FullException;

    /**
     * 自定义填充
     *
     * @param document    文件
     * @param fullDataMap 参数集合
     * @param handleType  处理类型
     */
    public void customizeFull(XWPFDocument document, FullDataMap fullDataMap, String handleType) throws FullException {
    }

    @Override
    public void close() {
    }
}
