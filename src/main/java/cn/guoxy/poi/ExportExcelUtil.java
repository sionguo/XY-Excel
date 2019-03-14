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

import java.io.ByteArrayOutputStream;
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

    private static Integer max_sheet = 1000;
    private static Integer max_row = 60000;

    /**
     * 导出Excel对象
     *
     * @param sheetDataListArr Excel数据
     * @return Workbook
     */
    public static Workbook exportWorkbook(List<?>... sheetDataListArr) {

        // data array valid
        if (sheetDataListArr == null || sheetDataListArr.length == 0) {
            logger.error("参数异常");
            throw new ParameterErrorException("参数异常");
        }

        // book （HSSFWorkbook=2003/xls、XSSFWorkbook=2007/xlsx）
        Workbook workbook = new HSSFWorkbook();

        // sheet
        for (List<?> dataList : sheetDataListArr) {
            makeSheet(workbook, dataList);
        }

        return workbook;
    }

    private static void makeSheet(Workbook workbook, List<?> sheetDataList) {
        // data
        if (sheetDataList == null || sheetDataList.size() == 0) {
            logger.error("参数异常");
            throw new ParameterErrorException("参数异常");
        }

        // sheet
        Class<?> sheetClass = sheetDataList.get(0).getClass();
        SheetSetting sheetSetting = sheetClass.getAnnotation(SheetSetting.class);

        String sheetName = sheetDataList.get(0).getClass().getSimpleName();
        int headColorIndex = -1;
        if (sheetSetting != null) {
            if (sheetSetting.name() != null && sheetSetting.name().trim().length() > 0) {
                sheetName = sheetSetting.name().trim();
            }
            headColorIndex = sheetSetting.headColor().getIndex();
        }

        Sheet existSheet = workbook.getSheet(sheetName);
        String newSheet =sheetName;
        if (existSheet != null) {
            for (int i = 2; i <= max_sheet; i++) {
                String newSheetName = sheetName.concat(String.valueOf(i));
                existSheet = workbook.getSheet(newSheetName);
                if (existSheet == null) {
                    newSheet = newSheetName;
                    break;
                } else {
                    continue;
                }
            }
        }

        Sheet sheet = workbook.createSheet(newSheet);

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
            logger.error("参数异常");
            throw new ParameterErrorException("参数异常");
        }

        // sheet header row
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
        int rowIndex = 0;
        for (int dataIndex = 0; dataIndex < sheetDataList.size(); dataIndex++) {
            rowIndex++;
            if (rowIndex >= max_row) {
                sheet = createSheet(workbook, sheetName);
                rowIndex = 0;
            }
            Object rowData = sheetDataList.get(dataIndex);

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

    private static Sheet createSheet(Workbook workbook, String sheetName) {
        Sheet existSheet = workbook.getSheet(sheetName);
        if (existSheet != null) {
            for (int i = 2; i <= max_sheet; i++) {
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

    /**
     * 导出Excel文件到磁盘
     *
     * @param filePath
     * @param sheetDataListArr 数据，可变参数，如多个参数则代表导出多张Sheet
     */
    public static void exportToFile(String filePath, List<?>... sheetDataListArr) {
        // workbook
        Workbook workbook = exportWorkbook(sheetDataListArr);

        FileOutputStream fileOutputStream = null;
        try {
            // workbook 2 FileOutputStream
            fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);

            // flush
            fileOutputStream.flush();
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
     * 导出Excel字节数据
     *
     * @param sheetDataListArr
     * @return byte[]
     */
    public static byte[] exportToBytes(List<?>... sheetDataListArr) {
        // workbook
        Workbook workbook = exportWorkbook(sheetDataListArr);

        ByteArrayOutputStream byteArrayOutputStream = null;
        byte[] result = null;
        try {
            // workbook 2 ByteArrayOutputStream
            byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.write(byteArrayOutputStream);

            // flush
            byteArrayOutputStream.flush();

            result = byteArrayOutputStream.toByteArray();
            return result;
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
            throw new XYExcelException(e);
        } finally {
            try {
                if (byteArrayOutputStream != null) {
                    byteArrayOutputStream.close();
                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
                throw new XYExcelException(e);
            }
        }
    }

}
