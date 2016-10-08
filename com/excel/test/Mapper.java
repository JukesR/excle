package com.excel.test;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by smt2 on 16-10-8.
 */
public class Mapper {
    private static final Logger logger = LoggerFactory.getLogger(Mapper.class);
    private static DataFormatter formatter = new DataFormatter();
    private List<Field> fieldList = null;
    private List<Column> columnList = null;
    private String excelName = "";
    public <T> File write(String path, List<T> beans) throws Exception {
        if (init(beans)) {
            return write(path, fieldList, columnList, beans);
        }
        return null;
    }


    private <T> boolean init(List<T> beans) {
        boolean initSuccess = false;
        this.fieldList = new ArrayList<>();
        this.columnList = new ArrayList<>();
        if (beans != null && !beans.isEmpty()) {
            T t = beans.get(0);
            if (t.getClass().getDeclaredAnnotation(Excel.class) != null) {
                this.excelName = t.getClass().getAnnotation(Excel.class).name();
                Field[] fields = t.getClass().getDeclaredFields();
                for (Field field : fields) {
                    Column column = field.getDeclaredAnnotation(Column.class);
                    if (column != null) {
                        this.fieldList.add(field);
                        this.columnList.add(column);
                        initSuccess = true;
                    }
                }
            }
            return initSuccess;
        }
        return false;
    }

    private List<T> read(File file, T t) throws Exception {
        if (file != null) {
            Map<String, List<String>> table = readToTable(file);

        }
        return null;
    }


    private Map<String, List<String>> readToTable(File file) throws Exception {
        if (file != null) {
            HSSFWorkbook xssfWorkbook = new HSSFWorkbook(new FileInputStream(file));
            Sheet sheet = xssfWorkbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            List<String> title = new ArrayList<String>();
            for (Cell cell : row) {
                title.add(getCellStringValue(cell));
            }
            Map<String, List<String>> table = initTable(title);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                row = sheet.getRow(rowIndex);
                for (int i = 0; i < title.size(); i++) {
                    table.get(title.get(i)).set(rowIndex - 1, getCellStringValue(row.getCell(i)));
                }
            }
            return table;
        }
        return null;
    }

    private Map<String, List<String>> initTable(List<String> title) {
        Map<String, List<String>> table = new HashMap<>();
        for (String s : title) {
            table.put(s, new ArrayList<>());

        }
        return table;
    }

    private <T> File write(String path, List<Field> fields, List<Column> columns, List<T> list) throws Exception {
        File file = FileUtils.getFile(path.toString() + excelName + ".xls");
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet1");
        Row row = sheet.createRow(0);
        sheet.setDefaultRowHeightInPoints((short) 15);
        HSSFFont headerFont = (HSSFFont) workbook.createFont();
        headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerFont.setFontHeightInPoints((short) 10);
        row.setHeight((short) (25 * 20));

        for (int i = 0; i < columns.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(columns.get(i).name());
            sheet.setColumnWidth(i, columns.get(i).width() * 100);
        }
        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.size(); j++) {
                Cell cell = row.createCell(j);
                if (columns.get(i).converter() == void.class) {
                    cell.setCellValue(BeanUtils.getProperty(list.get(i), fields.get(j).getName()));
                } else {
                    Convert convert = (Convert) columns.get(i).converter().newInstance();
                    cell.setCellValue(convert.beanToExcel(BeanUtils.getProperty(list.get(i), fields.get(j).getName())));
                }
            }
        }
        workbook.write(FileUtils.openOutputStream(file));
        return file;
    }


    private static String getCellStringValue(Cell cell) {

        String cellValue = "";
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                cellValue = cell.getStringCellValue();
                if (cellValue.trim().equals("") || cellValue.trim().length() <= 0)
                    cellValue = " ";
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
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
            case HSSFCell.CELL_TYPE_FORMULA:
                cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
//              value = String.valueOf(cell.getNumericCellValue());
                cellValue = getCellStringValue(cell);
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                cellValue = " ";
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                break;
            default:
                break;
        }
        return cellValue;
    }
}
