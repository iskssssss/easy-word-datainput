package top.kongsheng.common.word.datainput.core;

import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.input.AbsDataInput;
import top.kongsheng.common.word.datainput.model.input.BatchDataInput;
import top.kongsheng.common.word.datainput.model.input.ImageDataInput;
import top.kongsheng.common.word.datainput.model.input.TextDataInput;
import top.kongsheng.common.word.datainput.utils.FullDataInfoUtil;

import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * 填充参数集合
 *
 * @author 孔胜
 * @date 2023/9/1 20:15
 */
public class FullDataMap extends LinkedHashMap<String, AbsDataInput<?>> {

    public FullDataMap() throws FullException {
        this.set("now", LocalDateTime.now());
    }

    public <T extends AbsDataInput> List<T> get(Object key, Class<T > c)  {
        List<T> resultList = new LinkedList<>();
        AbsDataInput<?> fullDataInfo = super.get(key);
        if (fullDataInfo.getClass().isAssignableFrom(c)) {
            resultList.add(((T) fullDataInfo));
        } else {
            if (fullDataInfo instanceof BatchDataInput) {
                List<AbsDataInput<?>> list = ((BatchDataInput) fullDataInfo).getList();
                for (AbsDataInput<?> item : list) {
                    if (item.getClass().isAssignableFrom(c)) {
                        resultList.add(((T) item));
                    }
                }
            }
        }
        return resultList;
    }

    /**
     * 设置填充参数
     *
     * @param bean 值
     */
    public void setBean(Object bean) throws FullException {
        FullDataMap fullDataMap = toDataMap(bean);
        this.putAll(fullDataMap);
    }

    public ImageDataInput setImageBase64(String key, String imageBase64) throws FullException {
        return this.set(ImageDataInput.base64To(imageBase64).setKey(key));
    }

    /**
     * 设置填充参数
     *
     * @param fullDataInfo 值
     */
    public <T extends AbsDataInput<T>> T set(T fullDataInfo) throws FullException {
        String key = fullDataInfo.getKey();
        if (StrUtil.isEmpty(key)) {
            throw new FullException("填充参数添加失败，键不可为空。");
        }
        this.put(key, fullDataInfo);
        String extendKey = fullDataInfo.getExtendKey();
        if (StrUtil.isNotEmpty(extendKey)) {
            this.put(extendKey, fullDataInfo);
        }
        return fullDataInfo;
    }

    /**
     * 设置填充参数
     *
     * @param key   键
     * @param value 值
     * @return 结果
     */
    public TextDataInput set(String key, Object value) throws FullException {
        if (StrUtil.isEmpty(key)) {
            throw new FullException("填充参数添加失败，键不可为空。");
        }
        TextDataInput textDataInput = TextDataInput.to(value);
        this.put(key, textDataInput);
        return textDataInput;
    }

    /**
     * 设置填充参数
     *
     * @param key       键
     * @param valueList 值列表
     */
    public void set(String key, List<AbsDataInput<?>> valueList) throws FullException {
        if (StrUtil.isEmpty(key)) {
            throw new FullException("填充参数添加失败，键不可为空。");
        }
        if (valueList == null || valueList.isEmpty()) {
            super.remove(key);
            return;
        }
        this.put(key, BatchDataInput.to(valueList));
    }

    @Override
    public AbsDataInput put(String key, AbsDataInput fullDataInfo) {
        if (fullDataInfo == null) {
            super.remove(key);
            return null;
        }
        if (StrUtil.isEmpty(fullDataInfo.getKey())) {
            fullDataInfo.setKey(key);
        }
        return super.put(key, fullDataInfo);
    }

    /**
     * 添加填充参数
     *
     * @param fullDataInfo 值
     * @return 结果
     */
    public <T extends AbsDataInput<T>> T put(T fullDataInfo) throws FullException {
        this.insert(-1, fullDataInfo.getKey(), fullDataInfo, true);
        String extendKey = fullDataInfo.getExtendKey();
        if (StrUtil.isNotEmpty(extendKey)) {
            this.insert(-1, extendKey, fullDataInfo, false);
        }
        return fullDataInfo;
    }

    /**
     * 添加填充参数
     *
     * @param fullDataInfoList 值列表
     * @return 结果
     */
    public void put(List fullDataInfoList) throws FullException {
        for (Object fullDataInfo : fullDataInfoList) {
            if (fullDataInfo instanceof AbsDataInput) {
                this.put(((AbsDataInput) fullDataInfo));
            }
        }
    }

    /**
     * 添加填充参数
     *
     * @param key       键
     * @param valueList 值列表
     * @return 结果
     */
    public void put(String key, List<AbsDataInput<?>> valueList) {
        if (valueList == null || valueList.isEmpty()) {
            super.remove(key);
            return;
        }
        AbsDataInput<?> oldFullDataInfo = super.get(key);
        if (oldFullDataInfo instanceof BatchDataInput) {
            ((BatchDataInput) oldFullDataInfo).addAll(valueList);
            return;
        }
        BatchDataInput batchDataInput = BatchDataInput.to(oldFullDataInfo);
        batchDataInput.addAll(valueList);
        this.put(key, batchDataInput);
    }

    /**
     * 添加填充参数
     *
     * @param key   键
     * @param value 值
     * @return 结果
     */
    public TextDataInput putValue(String key, Object value) throws FullException {
        return this.insert(-1, key, value);
    }

    /**
     * 插入填充参数
     *
     * @param fullDataInfo 值
     * @return 结果
     */
    public <T extends AbsDataInput<T>> T insert(T fullDataInfo) throws FullException {
        return insert(0, fullDataInfo);
    }

    /**
     * 插入填充参数
     *
     * @param i            插入下标
     * @param fullDataInfo 值
     * @return 结果
     */
    public <T extends AbsDataInput<T>> T insert(int i, T fullDataInfo) throws FullException {
        this.insert(i, fullDataInfo.getKey(), fullDataInfo, true);
        return fullDataInfo;
    }

    /**
     * 插入填充参数
     *
     * @param key   键
     * @param value 值
     * @return 结果
     */
    public TextDataInput insert(String key, Object value) throws FullException {
        TextDataInput textDataInput = TextDataInput.to(value);
        this.insert(0, key, textDataInput, true);
        return textDataInput;
    }

    /**
     * 插入填充参数
     *
     * @param i     插入下标
     * @param key   键
     * @param value 值
     * @return 结果
     */
    public TextDataInput insert(int i, String key, Object value) throws FullException {
        TextDataInput textDataInput = TextDataInput.to(value);
        this.insert(i, key, textDataInput, true);
        return textDataInput;
    }

    /**
     * 插入填充参数
     *
     * @param i            插入下标
     * @param key          键
     * @param fullDataInfo 值
     * @param isThrow      是否抛出异常
     * @return 结果
     */
    public AbsDataInput<?> insert(int i, String key, AbsDataInput<?> fullDataInfo, boolean isThrow) throws FullException {
        if (StrUtil.isEmpty(key) && isThrow) {
            throw new FullException("填充参数添加失败，键不可为空。");
        }
        AbsDataInput<?> oldFullDataInfo = super.get(key);
        if (oldFullDataInfo == null) {
            this.put(key, fullDataInfo);
            return fullDataInfo;
        }
        if (oldFullDataInfo instanceof BatchDataInput) {
            ((BatchDataInput) oldFullDataInfo).add(i, fullDataInfo);
            return fullDataInfo;
        }
        BatchDataInput batchDataInput = BatchDataInput.to(oldFullDataInfo);
        batchDataInput.add(i, fullDataInfo);
        return this.put(key, batchDataInput);
    }

    @Override
    public void putAll(Map<? extends String, ? extends AbsDataInput<?>> m) {
        for (Map.Entry<? extends String, ? extends AbsDataInput<?>> entry : m.entrySet()) {
            this.put(entry.getKey(), entry.getValue());
        }
    }

    /**
     * 敏感内容替换
     *
     * @param keys 键
     */
    public void encryption(String... keys) {
        for (String key : keys) {
            this.put(key, TextDataInput.to("--"));
        }
    }

    /**
     * 切割内容
     *
     * @param keys 键
     */
    public void split(String... keys) {
        this.split(0, false, keys);
    }

    /**
     * 切割内容
     *
     * @param indentationFirstLine     首行缩进
     * @param autoIndentationFirstLine 是否自动判断是否需要首行缩进
     * @param keys                     键
     */
    public void split(int indentationFirstLine, boolean autoIndentationFirstLine, String... keys) {
        if (keys == null || keys.length < 1) {
            return;
        }
        for (String key : keys) {
            if (StrUtil.isEmpty(key)) {
                continue;
            }
            FullDataInfoUtil.splitValue(this, key, indentationFirstLine, autoIndentationFirstLine);
        }
    }

    public static <T> FullDataMap toDataMap(T t) throws FullException {
        return toDataMap(t, "");
    }

    public static <T> FullDataMap toDataMap(T t, String prefix) throws FullException {
        final Class<?> aClass = t.getClass();
        final Field[] declaredFields = aClass.getDeclaredFields();
        FullDataMap resultMap = new FullDataMap();
        String finalPrefix = StrUtil.isEmpty(prefix) ? "" : (prefix + ".");
        for (Field declaredField : declaredFields) {
            Object fieldValue = ReflectUtil.getFieldValue(t, declaredField);
            final String fieldName = declaredField.getName();
            String finalKey = finalPrefix + fieldName;
            if (resultMap.containsKey(finalKey)) {
                BatchDataInput batchDataInput = BatchDataInput.to(resultMap.get(finalKey)).setKey(finalKey);
                resultMap.put(finalKey, batchDataInput);
            }
            AbsDataInput<?> absDataInput = resultMap.get(finalKey);
            TextDataInput textDataInput = TextDataInput.to(fieldValue).setKey(finalKey);
            if (absDataInput instanceof BatchDataInput) {
                ((BatchDataInput) absDataInput).add(textDataInput);
                continue;
            }
            resultMap.put(finalKey, textDataInput);
        }
        return resultMap;
    }
}
