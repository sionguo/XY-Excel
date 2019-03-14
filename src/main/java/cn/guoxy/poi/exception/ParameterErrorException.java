package cn.guoxy.poi.exception;

/**
 * @author: XiaoyongGuo
 * @date: 2019/3/14/014 19:45
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
