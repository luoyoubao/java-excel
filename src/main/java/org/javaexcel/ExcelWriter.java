package org.javaexcel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
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
    protected DataFormat datafmt;

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
        // 设置表头样式
        ExcelHeader header = metedata.getHeader();
        if (metedata.isHasHeader() && null != header) {
            setCellStyle(header.getCellStyle());
            this.stylesMap.put("headerStyle", cellStyle);
        }

        // 设置footer样式
        ExcelFooter footer = metedata.getFooter();
        if (metedata.isHasFooter() && null != footer) {
            setCellStyle(footer.getCellStyle());
            this.stylesMap.put("footerStyle", cellStyle);
        }

        // 设置标题样式
        this.setCellStyle(metedata.getTitleStyle());
        stylesMap.put("titleStyle", cellStyle);

        // 设置数据单元格样式
        for (ExcelTitle excelTitle : allTitles) {
            this.setCellStyle(excelTitle.getCellStyle());
            setDataFormat(excelTitle);
            stylesMap.put("cellstyle_" + excelTitle.getIndex(), cellStyle);
        }
    }

    private void setCellStyle(ExcelCellStyle style) {
        cellStyle = wb.createCellStyle();
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        if (null != style) {
            cellStyle.setAlignment(style.getAlign());
            cellStyle.setVerticalAlignment(style.getVerticalAlign());
            font = wb.createFont();
            font.setFontHeightInPoints(style.getSize());
            font.setColor(style.getColor());
            font.setItalic(style.isItalic());
            cellStyle.setFont(font);
            cellStyle.setFillForegroundColor(style.getBackgroundColor());
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        }
        setBorder();
    }

    private void setDataFormat(ExcelTitle ete) {
        datafmt = wb.createDataFormat();
        switch (ete.getDataType()) {
            case NUMBER:
                this.cellStyle.setDataFormat(datafmt.getFormat("###.00"));
                break;
            case INT:
                this.cellStyle.setDataFormat(datafmt.getFormat("0"));
                break;
            case MONEY:
                this.cellStyle.setDataFormat(datafmt.getFormat("#,##0.00"));
                break;
            case PERCENT:
                this.cellStyle.setDataFormat(datafmt.getFormat("00.00%"));
                break;
            default:
                break;
        }
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