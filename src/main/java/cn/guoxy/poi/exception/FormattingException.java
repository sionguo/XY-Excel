package cn.guoxy.poi.exception;

/**
 * @author: XiaoyongGuo
 * @date: 2019/3/14/014 20:24
 */
public class FormattingException extends RuntimeException {
    public FormattingException() {
        super();
    }

    public FormattingException(String message) {
        super(message);
    }

    public FormattingException(String message, Throwable cause) {
        super(message, cause);
    }

    public FormattingException(Throwable cause) {
        super(cause);
    }
}
