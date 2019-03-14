package cn.guoxy.poi.exception;

/**
 * @author: XiaoyongGuo
 * @date: 2019/3/14/014 20:34
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
