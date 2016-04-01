package org.javaexcel.model;

import java.util.List;

/*
 * File name   : ExeclTitle.java
 * @Copyright  : www.quancheng-ec.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class ExcelTitle {
    // 列索引
    private Integer index;

    // 列名
    private String name;

    // 列标题
    private String displayName;

    // 列类型(number, double, date, money, text, percent)
    private CellType dataType;

    private CellStyle cellStyle;

    // 列是否支持合并(行合并)
    private boolean isMerge = false;

    // 是否有子标题(只支持一级子标题)
    private boolean hasSubTitle = false;

    // 子标题列表
    private List<ExcelTitle> subTitles;

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDisplayName() {
        return displayName;
    }

    public void setDisplayName(String displayName) {
        this.displayName = displayName;
    }

    public CellType getDataType() {
        return dataType;
    }

    public void setDataType(CellType dataType) {
        this.dataType = dataType;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public boolean isMerge() {
        return isMerge;
    }

    public void setMerge(boolean isMerge) {
        this.isMerge = isMerge;
    }

    public boolean isHasSubTitle() {
        return hasSubTitle;
    }

    public void setHasSubTitle(boolean hasSubTitle) {
        this.hasSubTitle = hasSubTitle;
    }

    public List<ExcelTitle> getSubTitles() {
        return subTitles;
    }

    public void setSubTitles(List<ExcelTitle> subTitles) {
        this.subTitles = subTitles;
    }
}