package cn.guoxy.poi.exception;

/**
 * @author XiaoyongGuo
 * @version 1.0-SNAPSHOT
 */
public class XYExcelException extends RuntimeException {
    public XYExcelException() {
    }

    public XYExcelException(String message) {
        super(message);
    }

    public XYExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public XYExcelException(Throwable cause) {
        super(cause);
    }
}
