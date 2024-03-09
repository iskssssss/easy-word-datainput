package top.kongsheng.common.word.datainput.model;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.lang.Pair;
import cn.hutool.core.util.StrUtil;
import lombok.Data;
import lombok.ToString;

import java.io.*;
import java.util.Arrays;

/**
 * 模板信息
 *
 * @author 孔胜
 * @date 2023/7/28 9:59
 */
@Data
@ToString
public class TemplateInfo implements Serializable {
    private static final long serialVersionUID = 42L;

    /**
     * 模板下载地址
     */
    private String url;
    /**
     * 处理类型
     */
    private String handleType;
    /**
     * 文件名称
     */
    private String fileName;
    /**
     * 文件后缀
     */
    private String fileSuffix;
    /**
     * 文件字节
     */
    private byte[] fileBytes;

    public void setFileName(String fileName) {
        this.fileName = fileName;
        this.fileSuffix = FileUtil.getSuffix(fileName);
    }

    public InputStream createInputStream() {
        int length = this.fileBytes.length;
        byte[] copyBytes = Arrays.copyOf(this.fileBytes, length);
        return new ByteArrayInputStream(copyBytes);
    }

    public Pair<String, InputStream> createResultPair() {
        return new Pair<>(this.getFileSuffix(), this.createInputStream());
    }

    public String getHandleType() {
        if (StrUtil.isEmpty(this.handleType)) {
            this.handleType = "DEFAULT";
        }
        return this.handleType;
    }

    public void setFile(File file) throws IOException {
        String name = file.getName();
        this.setFileName(name);
        this.fileBytes = IoUtil.readBytes(FileUtil.getInputStream(file));
    }
}
