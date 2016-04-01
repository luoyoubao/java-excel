package org.javaexcel.model;

/*
 * File name   : ExcelHeader.java
 * @Copyright  : www.quancheng-ec.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class ExcelHeader {
    // 表头名称
    private String headerName;

    // 样式
    private CellStyle cellStyle;

    public String getHeaderName() {
        return headerName;
    }

    public void setHeaderName(String headerName) {
        this.headerName = headerName;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}