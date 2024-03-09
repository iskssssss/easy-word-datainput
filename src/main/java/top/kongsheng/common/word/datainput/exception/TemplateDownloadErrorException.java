package top.kongsheng.common.word.datainput.exception;

/**
 * 导出模板下载错误异常
 *
 * @author 孔胜
 * @date 2023/9/3 10:55
 */
public class TemplateDownloadErrorException extends FullException {

    public TemplateDownloadErrorException(String message) {
        super(message);
    }

    public TemplateDownloadErrorException(String message, Throwable cause) {
        super(message, cause);
    }
}
