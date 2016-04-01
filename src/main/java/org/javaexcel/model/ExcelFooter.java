package org.javaexcel.model;

/*
 * File name   : ExcelFooter.java
 * @Copyright  : www.quancheng-ec.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class ExcelFooter {
    // EXCEL底部备注
    private String remarks;

    // 样式
    private CellStyle cellStyle;

    public String getRemarks() {
        return remarks;
    }

    public void setRemarks(String remarks) {
        this.remarks = remarks;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}