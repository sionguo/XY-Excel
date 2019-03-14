package cn.guoxy.poi.exception;

/**
 * @author XiaoyongGuo
 * @version 1.0-SNAPSHOT
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
