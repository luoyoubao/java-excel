package org.javaexcel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.javaexcel.model.ExcelCellStyle;
import org.javaexcel.model.ExcelFooter;
import org.javaexcel.model.ExcelHeader;
import org.javaexcel.model.ExcelMetaData;
import org.javaexcel.model.ExcelTitle;

/*
 * File name   : ExcelWriter.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public abstract class ExcelWriter {
    protected static final short DEFAULTROWHEIGHT = 16;
    // 样式表
    protected Map<String, CellStyle> stylesMap = new HashMap<String, CellStyle>();
    // 存储所有的标题(子级标题)
    protected List<ExcelTitle> allTitles = new ArrayList<ExcelTitle>();
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

    protected void init() throws Exception {
        List<ExcelTitle> titles = this.metedata.getExcelTitle();
        if (null == titles || titles.isEmpty()) {
            throw new Exception("The excel title is empty.");
        }
        for (ExcelTitle excelTitle : titles) {
            if (null != excelTitle.getSubTitles() && !excelTitle.getSubTitles().isEmpty()) {
                // 列设置不允许既需要合并单元格又有子标题
                if (excelTitle.isMerge()) {
                    throw new Exception("The column has sub title that cannot merge the cell.");
                }

                allTitles.addAll(excelTitle.getSubTitles());
                columnSize += excelTitle.getSubTitles().size();
                continue;
            }
            allTitles.add(excelTitle);
            columnSize++;
        }
    }

    protected void initStyle() {
        ExcelCellStyle style = null;
        ExcelHeader header = metedata.getHeader();
        if (metedata.isHasHeader() && null != header) {
            cellStyle = wb.createCellStyle();
            style = header.getCellStyle();
            if (null != style) {
                cellStyle.setAlignment(style.getAlign());
                cellStyle.setVerticalAlignment(style.getVerticalAlign());

                font = wb.createFont();
                font.setFontHeightInPoints(style.getSize());
                font.setColor(style.getColor());
                cellStyle.setFont(font);
                cellStyle.setFillForegroundColor(style.getBackgroundColor());
                cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            }
            setBorder();
            this.stylesMap.put("headerStyle", cellStyle);
        }

        ExcelFooter footer = metedata.getFooter();
        if (metedata.isHasFooter() && null != footer) {
            cellStyle = wb.createCellStyle();
            style = footer.getCellStyle();
            if (null != style) {
                cellStyle.setAlignment(style.getAlign());
                cellStyle.setVerticalAlignment(style.getVerticalAlign());

                font = wb.createFont();
                font.setFontHeightInPoints(style.getSize());
                font.setColor(style.getColor());
                font.setItalic(style.isItalic());
                cellStyle.setFont(font);
            }
            setBorder();
            this.stylesMap.put("footerStyle", cellStyle);
        }

        cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        font = wb.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        cellStyle.setFont(font);
        setBorder();
        stylesMap.put("titleStyle", cellStyle);

        for (ExcelTitle excelTitle : allTitles) {
            cellStyle = wb.createCellStyle();
            cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
            ExcelCellStyle estyle = excelTitle.getCellStyle();
            if (null != estyle) {
                cellStyle.setAlignment(estyle.getAlign());
                cellStyle.setVerticalAlignment(estyle.getVerticalAlign());

                font = wb.createFont();
                font.setFontHeightInPoints(estyle.getSize());
                font.setColor(estyle.getColor());
                cellStyle.setFont(font);
            }
            setBorder();
            stylesMap.put("cellstyle_" + excelTitle.getIndex(), cellStyle);
        }
    }

    public String getFormat(int index) {
        return BuiltinFormats.getBuiltinFormat(index);
    }

    /**
     * 获取样式
     * 
     * @param key
     * @return
     */
    protected CellStyle getStyle(String key) {
        return stylesMap.get(key);
    }

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