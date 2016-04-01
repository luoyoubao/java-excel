package org.javaexcel.model;

/*
 * File name   : CellStyle.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class CellStyle {
    // 字体大小
    private String fontSize;

    // 水平居中
    private boolean isAlignCenter;

    // 垂直居中
    private boolean isVerticalCenter;

    // 字体颜色
    private String fontColor;

    // 背景颜色
    private String backGroundColor;

    public String getFontSize() {
        return fontSize;
    }

    public void setFontSize(String fontSize) {
        this.fontSize = fontSize;
    }

    public boolean isAlignCenter() {
        return isAlignCenter;
    }

    public void setAlignCenter(boolean isAlignCenter) {
        this.isAlignCenter = isAlignCenter;
    }

    public boolean isVerticalCenter() {
        return isVerticalCenter;
    }

    public void setVerticalCenter(boolean isVerticalCenter) {
        this.isVerticalCenter = isVerticalCenter;
    }

    public String getFontColor() {
        return fontColor;
    }

    public void setFontColor(String fontColor) {
        this.fontColor = fontColor;
    }

    public String getBackGroundColor() {
        return backGroundColor;
    }

    public void setBackGroundColor(String backGroundColor) {
        this.backGroundColor = backGroundColor;
    }
}