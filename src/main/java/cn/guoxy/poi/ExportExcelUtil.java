package cn.guoxy.poi;

import cn.guoxy.poi.annotation.FieldSetting;
import cn.guoxy.poi.annotation.SheetSetting;
import cn.guoxy.poi.exception.ParameterErrorException;
import cn.guoxy.poi.exception.XYExcelException;
import cn.guoxy.poi.util.FieldReflectionUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.List;

/**
 * @author XiaoyongGuo
 * @version 1.0-SNAPSHOT
 */
public class ExportExcelUtil {
    private static Logger logger = LoggerFactory.getLogger(ExportExcelUtil.class);

    /**
     * 导出数据到文件
     *
     * @param dataList 数据集合
     * @param filePath 文件路径
     * @param fileName 文件名称，不带后缀
     * @author XiaoyongGuo
     * @version 1.0-SNAPSHOT
     **/
    public static void exportToFile(String filePath, String fileName, List<?>... dataList) {
        List<Workbook> workbooks = exportWorkbook(dataList);
        FileOutputStream fileOutputStream = null;
        try {
            for (int i = 0; i < workbooks.size(); i++) {
                String absolutePath = filePath + File.separator + fileName + i + ".xls";
                File file = new File(filePath);
                if (!file.exists()) {
                    file.mkdirs();
                    file.createNewFile();
                }
                file = new File(absolutePath);
                if (!file.exists()) {
                    file.createNewFile();
                }
                Workbook workbook = workbooks.get(i);
                fileOutputStream = new FileOutputStream(file);
                workbook.write(fileOutputStream);

            }
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
                throw new XYExcelException(e);
            }
        }

    }

    /**
     * @param dataList 数据集合
     * @return java.util.List workbook集合
     * @author XiaoyongGuo
     * @version 1.0-SNAPSHOT
     **/

    public static List<Workbook> exportWorkbook(List<?>... dataList) {
        if (dataList == null || dataList.length == 0) {
            logger.error("发生错误，参数为空");
            throw new ParameterErrorException("发生错误，未传入任何参数");
        }
        List<Workbook> tempList = new ArrayList<>();
        for (List<?> temp : dataList) {
            Workbook workbook = new HSSFWorkbook();
            makeSheet(workbook, temp);
            tempList.add(workbook);
        }
        return tempList;
    }

    /**
     * @param workbook 文档对象
     * @param dataList 数据集合
     * @return void
     * @author XiaoyongGuo
     * @version 1.0-SNAPSHOT
     **/
    private static void makeSheet(Workbook workbook, List<?> dataList) {
        if (dataList == null || dataList.size() == 0) {
            logger.error("发生错误，传入的参数集合为空");
            throw new ParameterErrorException("发生错误，传入的参数集合为空");
        }
        if (workbook == null) {
            logger.error("发生错误，未成功创建WorkBook");
            throw new ParameterErrorException("发生错误，未成功创建WorkBook");
        }
        Class<?> sheetClass = dataList.get(0).getClass();
        SheetSetting sheetSetting = sheetClass.getAnnotation(SheetSetting.class);

        String sheetName = sheetClass.getSimpleName();


        int headColorIndex = -1;
        if (sheetSetting != null) {
            if (sheetSetting.name() != null && sheetSetting.name().trim().length() > 0) {
                sheetName = sheetSetting.name().trim();
            }
            headColorIndex = sheetSetting.headColor().getIndex();
        }

        Sheet existSheet = workbook.getSheet(sheetName);
        if (existSheet != null) {
            for (int i = 2; i <= 1000; i++) {
                String newSheetName = sheetName.concat(String.valueOf(i));
                existSheet = workbook.getSheet(newSheetName);
                if (existSheet == null) {
                    sheetName = newSheetName;
                    break;
                } else {
                    continue;
                }
            }
        }
        Sheet sheet = workbook.createSheet(sheetName);

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

        CellStyle[] fieldDataStyleArr = new CellStyle[fields.size()];
        int[] fieldWidthArr = new int[fields.size()];
        Row headRow = sheet.createRow(0);
        for (int i = 0; i < fields.size(); i++) {

            // field
            Field field = fields.get(i);
            FieldSetting fieldSetting = field.getAnnotation(FieldSetting.class);

            String fieldName = field.getName();
            int fieldWidth = 0;
            HorizontalAlignment align = null;
            if (fieldSetting != null) {
                if (fieldSetting.name() != null && fieldSetting.name().trim().length() > 0) {
                    fieldName = fieldSetting.name().trim();
                }
                fieldWidth = fieldSetting.width();
                align = fieldSetting.align();
            }

            // field width
            fieldWidthArr[i] = fieldWidth;

            // head-style、field-data-style
            CellStyle fieldDataStyle = workbook.createCellStyle();
            if (align != null) {
                fieldDataStyle.setAlignment(align);
            }
            fieldDataStyleArr[i] = fieldDataStyle;

            CellStyle headStyle = workbook.createCellStyle();
            headStyle.cloneStyleFrom(fieldDataStyle);
            if (headColorIndex > -1) {
                headStyle.setFillForegroundColor((short) headColorIndex);
                headStyle.setFillBackgroundColor((short) headColorIndex);
                headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            // head-field data
            Cell cellX = headRow.createCell(i, CellType.STRING);
            cellX.setCellStyle(headStyle);
            cellX.setCellValue(String.valueOf(fieldName));
        }
        // sheet data rows
        for (int dataIndex = 0; dataIndex < dataList.size(); dataIndex++) {
            int rowIndex = dataIndex + 1;
            if (rowIndex >= 60000) {
                sheet = createSheet(workbook, sheetName);
            }
            Object rowData = dataList.get(dataIndex);
            Row rowX = sheet.createRow(rowIndex);

            for (int i = 0; i < fields.size(); i++) {
                Field field = fields.get(i);
                try {
                    field.setAccessible(true);
                    Object fieldValue = field.get(rowData);

                    String fieldValueString = FieldReflectionUtil.formatValue(field, fieldValue);

                    Cell cellX = rowX.createCell(i, CellType.STRING);
                    cellX.setCellValue(fieldValueString);
                    cellX.setCellStyle(fieldDataStyleArr[i]);
                } catch (IllegalAccessException e) {
                    logger.error(e.getMessage(), e);
                    throw new XYExcelException(e);
                }
            }
            // sheet finally
            for (int i = 0; i < fields.size(); i++) {
                int fieldWidth = fieldWidthArr[i];
                if (fieldWidth > 0) {
                    sheet.setColumnWidth(i, fieldWidth);
                } else {
                    sheet.autoSizeColumn((short) i);
                }
            }
        }

    }

    /**
     * @param sheetName sheet名称
     * @param workbook  文档对象
     * @return org.apache.poi.ss.usermodel.Sheet
     * @author XiaoyongGuo
     * @version 1.0-SNAPSHOT
     **/
    private static Sheet createSheet(Workbook workbook, String sheetName) {
        Sheet existSheet = workbook.getSheet(sheetName);
        if (existSheet != null) {
            for (int i = 2; i <= 1000; i++) {
                String newSheetName = sheetName.concat(String.valueOf(i));
                existSheet = workbook.getSheet(newSheetName);
                if (existSheet == null) {
                    sheetName = newSheetName;
                    break;
                } else {
                    continue;
                }
            }
        }
        Sheet sheet = workbook.createSheet(sheetName);
        return sheet;
    }

}
