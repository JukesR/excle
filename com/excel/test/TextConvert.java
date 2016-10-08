package com.excel.test;


/**
 * Created by smt2 on 16-10-8.
 */
public class TextConvert implements Convert<String, String> {
    @Override
    public String beanToExcel(String s) {
        return s+"123";
    }

    @Override
    public String ExcelToBean(String e) {
        return null;
    }
}
