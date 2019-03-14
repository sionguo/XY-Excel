package cn.guoxy.poi.exception;

/**
 * @author XiaoyongGuo
 * @version 1.0-SNAPSHOT
 */
public class ParameterErrorException extends RuntimeException {
    public ParameterErrorException() {
    }

    public ParameterErrorException(Throwable cause) {
        super(cause);
    }

    public ParameterErrorException(String message) {
        super(message);
    }

    public ParameterErrorException(String message, Throwable cause) {
        super(message, cause);
    }
}
