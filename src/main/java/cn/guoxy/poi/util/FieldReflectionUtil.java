package cn.guoxy.poi.util;

import cn.guoxy.poi.annotation.FieldSetting;
import cn.guoxy.poi.exception.FormattingException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author XiaoyongGuo
 * @version 1.0-SNAPSHOT
 */
public final class FieldReflectionUtil {

    private static Logger logger = LoggerFactory.getLogger(FieldReflectionUtil.class);

    private FieldReflectionUtil() {
    }

    public static Byte parseByte(String value) {
        try {
            value = value.replaceAll("　", "");
            return Byte.valueOf(value);
        } catch (NumberFormatException e) {
            logger.error("非法输入，输入的值为:{}", value, e);
            throw new FormattingException("非法输入，" + value, e);
        }
    }

    public static Boolean parseBoolean(String value) {
        value = value.replaceAll("　", "");
        if (Boolean.TRUE.toString().equalsIgnoreCase(value)) {
            return Boolean.TRUE;
        } else if (Boolean.FALSE.toString().equalsIgnoreCase(value)) {
            return Boolean.FALSE;
        } else {
            logger.error("非法输入，输入的值为:{}", value);
            throw new FormattingException("非法输入，" + value);
        }
    }

    public static Integer parseInt(String value) {
        try {
            value = value.replaceAll("　", "");
            return Integer.valueOf(value);
        } catch (NumberFormatException e) {
            logger.error("非法输入，输入的值为:{}", value, e);
            throw new FormattingException("非法输入，" + value, e);
        }
    }

    public static Short parseShort(String value) {
        try {
            value = value.replaceAll("　", "");
            return Short.valueOf(value);
        } catch (NumberFormatException e) {
            logger.error("非法输入，输入的值为:{}", value, e);
            throw new FormattingException("非法输入，" + value, e);
        }
    }

    public static Long parseLong(String value) {
        try {
            value = value.replaceAll("　", "");
            return Long.valueOf(value);
        } catch (NumberFormatException e) {
            logger.error("非法输入，输入的值为:{}", value, e);
            throw new FormattingException("非法输入，" + value, e);
        }
    }

    public static Float parseFloat(String value) {
        try {
            value = value.replaceAll("　", "");
            return Float.valueOf(value);
        } catch (NumberFormatException e) {
            logger.error("非法输入，输入的值为:{}", value, e);
            throw new FormattingException("非法输入，" + value, e);
        }
    }

    public static Double parseDouble(String value) {
        try {
            value = value.replaceAll("　", "");
            return Double.valueOf(value);
        } catch (NumberFormatException e) {
            logger.error("非法输入，输入的值为:{}", value, e);
            throw new FormattingException("非法输入，" + value, e);
        }
    }

    public static Date parseDate(String value, FieldSetting FieldSetting) {
        try {
            String datePattern = "yyyy-MM-dd HH:mm:ss";
            if (FieldSetting != null) {
                datePattern = FieldSetting.dateformat();
            }
            SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
            return dateFormat.parse(value);
        } catch (ParseException e) {
            logger.error("非法输入，输入的值为:{}", value, e);
            throw new FormattingException("非法输入，" + value, e);
        }
    }

    /**
     * 参数解析 （支持：Byte、Boolean、String、Short、Integer、Long、Float、Double、Date）
     *
     * @param field 字段
     * @param value 值
     * @return Object
     * @author XiaoyongGuo
     * @version 1.0-SNAPSHOT
     */
    public static Object parseValue(Field field, String value) {
        Class<?> fieldType = field.getType();

        FieldSetting FieldSetting = field.getAnnotation(FieldSetting.class);
        if (value == null || value.trim().length() == 0) {
            return null;
        }
        value = value.trim();

        if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
            return parseBoolean(value);
        } else if (String.class.equals(fieldType)) {
            return value;
        } else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
            return parseShort(value);
        } else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
            return parseInt(value);
        } else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
            return parseLong(value);
        } else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
            return parseFloat(value);
        } else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
            return parseDouble(value);
        } else if (Date.class.equals(fieldType)) {
            return parseDate(value, FieldSetting);

        } else {

            logger.error("非法请求，输入的参数类型为包装类型，您输入的类型为:{}", fieldType);
            throw new FormattingException("非法请求，输入的参数类型为包装类型，您输入的类型为" + fieldType);
        }
    }

    /**
     * 参数格式化为String
     *
     * @param field 字段
     * @param value 值
     * @return String
     * @author XiaoyongGuo
     * @version 1.0-SNAPSHOT
     */
    public static String formatValue(Field field, Object value) {
        Class<?> fieldType = field.getType();

        FieldSetting fieldSetting = field.getAnnotation(FieldSetting.class);
        if (value == null) {
            return null;
        }

        if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (String.class.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Date.class.equals(fieldType)) {
            String datePattern = "yyyy-MM-dd HH:mm:ss";
            if (fieldSetting != null && fieldSetting.dateformat() != null) {
                datePattern = fieldSetting.dateformat();
            }
            SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
            return dateFormat.format(value);
        } else {
            logger.error("非法请求，输入的参数类型为包装类型，您输入的类型为:{}", fieldType);
            throw new FormattingException("非法请求，输入的参数类型为包装类型，您输入的类型为" + fieldType);
        }
    }

}
