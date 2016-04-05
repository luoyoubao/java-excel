package org.javaexcel.model;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

/*
 * File name   : CellStyle.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class ExcelCellStyle {
    // 水平居中
    private short align = XSSFCellStyle.ALIGN_LEFT;

    // 垂直居中
    private short verticalAlign = XSSFCellStyle.VERTICAL_CENTER;

    // 字体大小
    private short size = XSSFFont.DEFAULT_FONT_SIZE;

    // 字体颜色
    private short color = ExcelColor.DEFAULT_COLOR;

    public short getSize() {
        return size;
    }

    public short getAlign() {
        return align;
    }

    public void setAlign(short align) {
        this.align = align;
    }

    public short getVerticalAlign() {
        return verticalAlign;
    }

    public void setVerticalAlign(short verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    public void setSize(short size) {
        this.size = size;
    }

    public short getColor() {
        return color;
    }

    public void setColor(short color) {
        this.color = color;
    }
}