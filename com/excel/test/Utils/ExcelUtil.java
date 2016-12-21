package cn.com.singlemountaintech.dxmanagement.common.utils;

import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.beanutils.converters.BigDecimalConverter;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by smt2 on 16-6-6.
 */
public class ExcelUtil {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    private static DataFormatter formatter = new DataFormatter();

    static {
        ConvertUtils.register(new BigDecimalConverter("0"), BigDecimal.class);
    }
    public static <T> List<T> read(File file, List<String> fields, List<String> columns, int offset, Class<T> cls) throws Exception {
        if (CollectionUtils.isNotEmpty(fields) && CollectionUtils.isNotEmpty(columns) && fields.size() == columns.size()) {
            Table table = readToTable(file, offset);

            List<T> list = new ArrayList<T>();
            if (CollectionUtils.isNotEmpty(fields)
                    && CollectionUtils.isNotEmpty(columns)
                    && fields.size() == columns.size()
                    && table != null
                    && CollectionUtils.isNotEmpty(table.cellSet())) {
                for (int index = 1; index <= table.rowKeySet().size(); index++) {
                    T t = cls.newInstance();
                    for (int i = 0; i < fields.size(); i++) {
                        instantiateNestedProperties(t, fields.get(i));
                        BeanUtils.setProperty(t, fields.get(i), table.get(index, columns.get(i)));
                    }
                    list.add(t);
                }
                return list;
            }
        }
        return null;
    }

    public static <T> File write(Path path, List<String> fields, List<String> columns, List<T> list)throws Exception {
        return write(path, "sheet1", fields, columns, null, list);
    }

    public static <T> File write(Path path, List<String> fields, List<String> columns,List<String> formaters, List<T> list)throws Exception {
        return write(path, "sheet1", fields, columns, null, list);
    }



    public static <T> File write(Path path, String sheetName, List<String> fields, List<String> columns, List<String> formaters, List<T> list) throws Exception {
        File file = FileUtils.getFile(path.toString());
        Workbook workbook = createWorkbook(path);
        Sheet sheet = null;
        if (StringUtils.isNotEmpty(sheetName)) {
            sheet = workbook.createSheet(sheetName);
        } else {
            sheet = workbook.createSheet("sheet1");
        }
        Row row = sheet.createRow(0);
        sheet.setDefaultRowHeightInPoints((short) 15);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
        Font headerFont = workbook.createFont();
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setFontHeightInPoints((short) 10);
        headerStyle.setFont(headerFont);
        row.setHeight((short) (25 * 20));
        for (int i = 0; i < columns.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(columns.get(i));
        }
        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i + 1);
            rowFormat(workbook, row, fields, formaters, list.get(i));
        }
        workbook.write(FileUtils.openOutputStream(file));

        return file;
    }

    public static <T> File write(Path path, String sheetName, List<String> fields, List<String> columns, List<T> list) {
        HSSFWorkbook workbook = null;
        File file = FileUtils.getFile(path.toString());
        try {
            workbook = new HSSFWorkbook(new FileInputStream(file));
            HSSFSheet sheet = workbook.createSheet(sheetName);
            Row row = sheet.createRow(0);
            sheet.setDefaultRowHeightInPoints((short) 15);
            HSSFCellStyle headerStyle = (HSSFCellStyle) workbook.createCellStyle();
            headerStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            HSSFFont headerFont = (HSSFFont) workbook.createFont();
            headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            headerFont.setFontHeightInPoints((short) 10);
            headerStyle.setFont(headerFont);
            row.setHeight((short) (25 * 20));
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(headerStyle);
                cell.setCellValue(columns.get(i));
            }
            for (int i = 0; i < list.size(); i++) {
                row = sheet.createRow(i + 1);
                for (int j = 0; j < fields.size(); j++) {

                    row.createCell(j).setCellValue(BeanUtils.getProperty(list.get(i), fields.get(j)));
                }
            }
            workbook.write(FileUtils.openOutputStream(file));
        } catch (Exception e) {
            if (logger.isErrorEnabled()) {
                logger.error("writeExcel", e);
            }

        } finally {
            try {
                workbook.close();
            } catch (Exception e) {
                if (logger.isErrorEnabled()) {
                    logger.error("writeExcel", e);
                }
            }
        }
        return file;
    }

    public static Table readToTable(File file, int offset) throws Exception {
        if (file != null) {
            Table<Integer, String, String> table = HashBasedTable.create();
            Workbook workbook = getWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(offset);
            List<String> title = new ArrayList<String>();
            for (Cell cell : row) {
                title.add(getCellStringValue(cell));
            }
            for (int rowIndex = offset + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                row = sheet.getRow(rowIndex);
                if (row!=null){
                    for (int i = 0; i < title.size(); i++) {
                        table.put(rowIndex, title.get(i), getCellStringValue(row.getCell(i)));
                    }
                }
            }
            return table;
        }
        return null;
    }

    private static Workbook getWorkbook(File file) throws IOException {
        Workbook wb = null;
        if (file.getName().endsWith(EXCEL_XLS)) {     //Excel 2003
            wb = new HSSFWorkbook(new FileInputStream(file));
        } else if (file.getName().endsWith(EXCEL_XLSX)) {    // Excel 2007/2010
            wb = new XSSFWorkbook(new FileInputStream(file));
        }
        return wb;
    }
    private static Workbook createWorkbook(Path path) {
        Workbook wb = null;
        if (path.toString().endsWith(EXCEL_XLS)) {     //Excel 2003
            wb = new HSSFWorkbook();
        } else if (path.toString().endsWith(EXCEL_XLSX)) {    // Excel 2007/2010
            wb = new XSSFWorkbook();
        }
        return wb;
    }
    private static void instantiateNestedProperties(Object obj, String fieldName) {
        try {
            String[] fieldNames = fieldName.split("\\.");
            if (fieldNames.length > 1) {
                StringBuffer nestedProperty = new StringBuffer();
                for (int i = 0; i < fieldNames.length - 1; i++) {
                    String fn = fieldNames[i];
                    if (i != 0) {
                        nestedProperty.append(".");
                    }
                    nestedProperty.append(fn);

                    Object value = PropertyUtils.getProperty(obj, nestedProperty.toString());

                    if (value == null) {
                        PropertyDescriptor propertyDescriptor = PropertyUtils.getPropertyDescriptor(obj, nestedProperty.toString());
                        Class<?> propertyType = propertyDescriptor.getPropertyType();
                        Object newInstance = propertyType.newInstance();
                        PropertyUtils.setProperty(obj, nestedProperty.toString(), newInstance);
                    }
                }
            }
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        } catch (InvocationTargetException e) {
            throw new RuntimeException(e);
        } catch (NoSuchMethodException e) {
            throw new RuntimeException(e);
        } catch (InstantiationException e) {
            throw new RuntimeException(e);
        }
    }

/*    public static File format(File file, List<String> columns, List<String> formater, int offset) throws IOException {
        Workbook wb = getWorkbook(file);
        Sheet sheet = wb.getSheetAt(0);
        DataFormat format = wb.createDataFormat();
        List<CellStyle> cellStyles = new ArrayList<>();
        for (String s : formater) {
            if (StringUtils.isNotEmpty(s)) {
                CellStyle cellStyle = wb.createCellStyle();
                cellStyle.setDataFormat(format.getFormat(s));
                cellStyles.add(cellStyle);
            } else {
                cellStyles.add(wb.createCellStyle());
            }
        }
        Row row = null;
        for (int rowIndex = offset + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int i = 0; i < columns.size(); i++) {
                    row.getCell(i).setCellStyle(cellStyles.get(i));
                }
            }
        }
        wb.write(FileUtils.openOutputStream(file));
        return file;
    }*/

    public static <T> Row rowFormat(Workbook wb, Row row, List<String> fields, List<String> formater, T t) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        if (CollectionUtils.isNotEmpty(formater)) {
            List<CellStyle> cellStyles = getCellStyles(wb, formater);
            for (int j = 0; j < fields.size(); j++) {
                Object value = PropertyUtils.getProperty(t, fields.get(j));
                try {
                    setCellValue(row.createCell(j), value, cellStyles.get(j));
                } catch (Exception e) {
                    if (logger.isErrorEnabled()) {
                        logger.error("write", e);
                    }
                }
            }
        } else {
            for (int j = 0; j < fields.size(); j++) {
                String value = BeanUtils.getProperty(t, fields.get(j));
                try {
                    row.createCell(j).setCellValue(value);
                } catch (Exception e) {
                    if (logger.isErrorEnabled()) {
                        logger.error("write", e);
                    }
                }
            }
        }
        return row;
    }

    private static void setCellValue(Cell cell, Object value, CellStyle style) {
        if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if ((value instanceof BigDecimal) || (value instanceof Float) || (value instanceof Double)) {
            cell.setCellValue(Double.parseDouble(value.toString()));
        } else {
            cell.setCellValue(value.toString());
        }
        cell.setCellStyle(style);
    }

    private static List<CellStyle> getCellStyles(Workbook wb, List<String> formater) {
        DataFormat format = wb.createDataFormat();
        List<CellStyle> cellStyles = new ArrayList<>();
        for (String s : formater) {
            if (StringUtils.isNotEmpty(s)) {
                CellStyle cellStyle = wb.createCellStyle();
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setDataFormat(format.getFormat(s));
                cellStyles.add(cellStyle);
            } else {
                CellStyle cellStyle = wb.createCellStyle();
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyles.add(cellStyle);
            }
        }
        return cellStyles;
    }

    private static String getCellStringValue(Cell cell) {

        String cellValue = "";
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                cellValue = cell.getStringCellValue();
                if (cellValue.trim().equals("") || cellValue.trim().length() <= 0)
                    cellValue = " ";
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                    try {
                        cellValue = sdf.format(date);
                    } catch (Exception e) {
                        if (logger.isErrorEnabled()) {
                            logger.error("date parse exception", e);
                        }
                    }
                } else {
                    cellValue = String.valueOf(formatter.formatCellValue(cell));
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                Workbook wb = cell.getSheet().getWorkbook();
                CreationHelper crateHelper = wb.getCreationHelper();
                FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
                cellValue = getCellStringValue(evaluator.evaluateInCell(cell));
                break;
            case Cell.CELL_TYPE_BLANK:
                cellValue = "";
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                break;
            case Cell.CELL_TYPE_ERROR:
                break;
            default:
                break;
        }
        return cellValue;
    }
}
