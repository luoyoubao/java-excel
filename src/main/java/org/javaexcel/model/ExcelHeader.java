package org.javaexcel.model;

/*
 * File name   : ExcelHeader.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class ExcelHeader {
    // 表头名称
    private String headerName;

    // 样式
    private ExcelCellStyle cellStyle = new ExcelCellStyle();

    private float rowHeight = 50;

    public float getRowHeight() {
        return rowHeight;
    }

    public void setRowHeight(float rowHeight) {
        this.rowHeight = rowHeight;
    }

    public String getHeaderName() {
        return headerName;
    }

    public void setHeaderName(String headerName) {
        this.headerName = headerName;
    }

    public ExcelCellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(ExcelCellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}