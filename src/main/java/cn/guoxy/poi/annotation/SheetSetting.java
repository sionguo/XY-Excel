package cn.guoxy.poi.annotation;

import org.apache.poi.hssf.util.HSSFColor;

import java.lang.annotation.*;

/**
 * @author: XiaoyongGuo
 * @date: 2019/3/14/014 19:21
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface SheetSetting {
    String name() default "";

    HSSFColor.HSSFColorPredefined headColor() default HSSFColor.HSSFColorPredefined.LIGHT_BLUE;
}
