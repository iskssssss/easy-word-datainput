package top.kongsheng.common.word.datainput.core;

import cn.hutool.core.io.IoUtil;

import java.io.*;

/**
 * CloneByteArrayInputStream
 *
 * @author 孔胜
 * @date 2023/1/4 9:24
 */
public class CloneByteArrayInputStream implements Closeable {
    private final ByteArrayOutputStream outputStream;
    private byte[] bytes;

    public CloneByteArrayInputStream(InputStream input) {
        outputStream = new ByteArrayOutputStream();
        IoUtil.copy(input, outputStream);
        bytes = outputStream.toByteArray();
    }

    public CloneByteArrayInputStream(byte[] bytes) throws IOException {
        this.bytes = bytes;
        this.outputStream = new ByteArrayOutputStream();
        this.outputStream.write(bytes);
    }

    public InputStream get() {
        return new ByteArrayInputStream(bytes);
    }

    public boolean isEmpty() {
        return bytes == null || bytes.length < 1;
    }

    public ByteArrayOutputStream getOutputStream() {
        return outputStream;
    }

    @Override
    public void close() throws IOException {
        outputStream.close();
        bytes = null;
    }
}
