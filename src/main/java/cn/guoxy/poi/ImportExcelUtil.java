package cn.guoxy.poi;

import cn.guoxy.poi.annotation.SheetSetting;
import cn.guoxy.poi.exception.XYExcelException;
import cn.guoxy.poi.util.FieldReflectionUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author: XiaoyongGuo
 * @date: 2019/3/14/014 20:42
 */
public class ImportExcelUtil {
    private static Logger logger = LoggerFactory.getLogger(ImportExcelUtil.class);

    public static List<Object> importExcel(Workbook workbook, Class<?> sheetClass) {
        List<Object> sheetDataList = importSheet(workbook, sheetClass);
        return sheetDataList;
    }

    public static List<Object> importSheet(Workbook workbook, Class<?> sheetClass) {
        try {
            // sheet
            SheetSetting sheetSetting = sheetClass.getAnnotation(SheetSetting.class);
            String sheetName = (sheetSetting != null && sheetSetting.name() != null && sheetSetting.name().trim().length() > 0) ? sheetSetting.name().trim() : sheetClass.getSimpleName();

            // sheet field
            List<Field> fields = new ArrayList<Field>();
            if (sheetClass.getDeclaredFields() != null && sheetClass.getDeclaredFields().length > 0) {
                for (Field field : sheetClass.getDeclaredFields()) {
                    if (Modifier.isStatic(field.getModifiers())) {
                        continue;
                    }
                    fields.add(field);
                }
            }

            if (fields == null || fields.size() == 0) {
                logger.error("发生错误，数据字段不能为空");
                throw new XYExcelException("发生错误，数据字段不能为空");
            }
            List<Object> dataList = new ArrayList<Object>();
            // sheet data
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                if (sheet == null) {
                    return null;
                }
                Iterator<Row> rowIterator = sheet.rowIterator();
                int rowIndex = 0;
                while (rowIterator.hasNext()) {
                    Row rowX = rowIterator.next();
                    if (rowIndex > 0) {
                        Object rowObj = sheetClass.newInstance();
                        for (int i = 0; i < fields.size(); i++) {
                            // cell
                            Cell cell = rowX.getCell(i);
                            if (cell == null) {
                                continue;
                            }
                            // call val str
                            cell.setCellType(CellType.STRING);
                            String fieldValueStr = cell.getStringCellValue();
                            // java val
                            Field field = fields.get(i);
                            Object fieldValue = FieldReflectionUtil.parseValue(field, fieldValueStr);
                            if (fieldValue == null) {
                                continue;
                            }
                            // fill val
                            field.setAccessible(true);
                            field.set(rowObj, fieldValue);
                        }
                        dataList.add(rowObj);
                    }
                    rowIndex++;
                }
            }

            return dataList;
        } catch (IllegalAccessException e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        } catch (InstantiationException e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        }
    }

    /**
     * 导入Excel文件，并封装成对象
     *
     * @param excelFile
     * @param sheetClass
     * @return List<Object>
     */
    public static List<Object> importExcel(File excelFile, Class<?> sheetClass) {
        try {
            Workbook workbook = WorkbookFactory.create(excelFile);
            List<Object> dataList = importExcel(workbook, sheetClass);
            return dataList;
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        } catch (InvalidFormatException e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        }
    }


    /**
     * 从文件路径导入Excel文件，并封装成对象
     *
     * @param filePath
     * @param sheetClass
     * @return List<Object>
     */
    public static List<Object> importExcel(String filePath, Class<?> sheetClass) {
        File excelFile = new File(filePath);
        List<Object> dataList = importExcel(excelFile, sheetClass);
        return dataList;
    }

    /**
     * 导入Excel数据流，并封装成对象
     *
     * @param inputStream
     * @param sheetClass
     * @return List<Object>
     */
    public static List<Object> importExcel(InputStream inputStream, Class<?> sheetClass) {
        try {
            Workbook workbook = WorkbookFactory.create(inputStream);
            List<Object> dataList = importExcel(workbook, sheetClass);
            return dataList;
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        } catch (InvalidFormatException e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        }
    }

}
