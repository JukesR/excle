package com.excel.test;

/**
 * Created by smt2 on 16-10-8.
 */
public interface Convert<E, B> {

    String beanToExcel(B b);

    B ExcelToBean(String e);
}
