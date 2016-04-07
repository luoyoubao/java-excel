package org.javaexcel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.javaexcel.model.ExcelMetaData;

/*
 * File name   : ExcelWriter.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public abstract class ExcelWriter {
    // 样式表
    protected Map<String, CellStyle> stylesMap = new HashMap<String, CellStyle>();
    protected List<Object> allDatas = new ArrayList<Object>();
    protected int rownum = 0;
    protected int columnSize = 0;
    protected Workbook wb;
    protected Sheet sheet;
    protected CellStyle cellStyle;
    protected Font font;
    protected Row row;
    protected Cell cell;
    protected ExcelMetaData metedata;

    /**
     * 导出数据到EXCEL
     * 
     * @param metadata
     * @param datas
     * @param filePath
     * @return
     */
    public abstract boolean process(ExcelMetaData metedata, List<Object> datas, String fileName) throws Exception;

    /**
     * 设置边框
     */
    protected void setBorder() {
        if (null != this.cellStyle) {
            cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
            cellStyle.setBorderTop(CellStyle.BORDER_THIN);
            cellStyle.setBorderRight(CellStyle.BORDER_THIN);
            cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        }
    }
}