package org.javaexcel.model;

/*
 * File name   : ExcelFooter.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class ExcelFooter {
    // EXCEL底部备注
    private String remarks;

    // 样式
    private ExcelCellStyle cellStyle;

    public String getRemarks() {
        return remarks;
    }

    public void setRemarks(String remarks) {
        this.remarks = remarks;
    }

    public ExcelCellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(ExcelCellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}