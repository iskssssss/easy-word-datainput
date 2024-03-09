package top.kongsheng.common.word.datainput.exception;

/**
 * 表达式异常
 *
 * @author 孔胜
 * @date 2023/9/8 14:40
 */
public class ExpressionFullException extends FullException {

    public ExpressionFullException(String message) {
        super(message);
    }

    public ExpressionFullException(String message, Throwable cause) {
        super(message, cause);
    }

    public static ExpressionFullException unsupportable(String condition) {
        return new ExpressionFullException("无法支持的表达式（" + condition + "）。");
    }
}
