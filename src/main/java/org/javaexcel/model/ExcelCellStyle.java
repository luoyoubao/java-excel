package org.javaexcel.model;

import org.apache.poi.ss.usermodel.CellStyle;

/*
 * File name   : CellStyle.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class ExcelCellStyle {
    // 水平居中
    private short align = CellStyle.ALIGN_LEFT;

    // 垂直居中
    private short verticalAlign = CellStyle.VERTICAL_CENTER;

    // 字体大小
    private short size = 11;

    // 字体颜色
    private short color = ExcelColor.DEFAULT_COLOR;

    // 单元格背景色
    private short backgroundColor = ExcelColor.GREY_25_PERCENT;

    // 斜体
    private boolean isItalic = false;

    public boolean isItalic() {
        return isItalic;
    }

    public void setItalic(boolean isItalic) {
        this.isItalic = isItalic;
    }

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

    public short getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(short backgroundColor) {
        this.backgroundColor = backgroundColor;
    }
}