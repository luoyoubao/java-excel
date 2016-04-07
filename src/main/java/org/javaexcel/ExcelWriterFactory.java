package org.javaexcel;

import org.javaexcel.model.ExcelType;
import org.javaexcel.xls.DataToExcelWriter;
import org.javaexcel.xlsx.XmlToExcelWriter;

/*
 * File name   : ExcelWriterFactory.java
 * @Copyright  : www.quancheng-ec.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月7日
 */
public class ExcelWriterFactory {
    public static ExcelWriter getWriter(String type) {
        ExcelType excelType = ExcelType.valueOf(type.toUpperCase());
        switch (excelType) {
            case XLS:
                return new DataToExcelWriter();
            case XLSX:
                return new XmlToExcelWriter();
            default:
                return null;
        }
    }
}