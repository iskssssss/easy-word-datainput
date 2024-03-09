package top.kongsheng.common.word.datainput.exception;

/**
 * 填充异常
 *
 * @author 孔胜
 * @date 2023/8/24 15:41
 */
public class FullException extends Exception {

    public FullException(String message) {
        super(message);
    }

    public FullException(String message, Throwable cause) {
        super(message, cause);
    }

}
