package org.javaexcel.model;

import org.apache.poi.hssf.util.CellReference;

/*
 * File name   : MergeCell.java
 * @Copyright  : www.quancheng-ec.com
 * Description : excelservice-service
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class CellMerge {
    private int beginColumn;
    private int endColumn;
    private int beginCell;
    private int endCell;

    public CellMerge(int beginColumn, int beginCell, int endColumn, int endCell) {
        this.beginColumn = beginColumn;
        this.beginCell = beginCell;
        this.endColumn = endColumn;
        this.endCell = endCell;
    }

    @Override
    public String toString() {
        StringBuffer sb = new StringBuffer();
        sb.append("<mergeCell ref=\"");
        sb.append(CellReference.convertNumToColString(beginCell)).append(beginColumn + 1);
        sb.append(":");
        sb.append(CellReference.convertNumToColString(endCell)).append(endColumn + 1);
        sb.append("\"/>");
        return sb.toString();
    }

    public int getBeginColumn() {
        return beginColumn;
    }

    public void setBeginColumn(int beginColumn) {
        this.beginColumn = beginColumn;
    }

    public int getEndColumn() {
        return endColumn;
    }

    public void setEndColumn(int endColumn) {
        this.endColumn = endColumn;
    }

    public int getBeginCell() {
        return beginCell;
    }

    public void setBeginCell(int beginCell) {
        this.beginCell = beginCell;
    }

    public int getEndCell() {
        return endCell;
    }

    public void setEndCell(int endCell) {
        this.endCell = endCell;
    }
}