package org.javaexcel.model;

import java.util.List;

/*
 * File name   : ExeclInfo.java
 * @Copyright  : luoyoub@163.com
 * Description : execlservice-common
 * Author      : Robert
 * CreateTime  : 2016年3月14日
 */
public class ExcelMetaData {
    // 文件类型(支持xls和xlsx)
    private String fileType = "xlsx";

    // Excel文件名称
    private String fileName = "service";

    // Excel sheet名称
    private String sheetName = "service";

    // 是否有大表头
    private boolean hasHeader = false;

    // 大表头
    private ExcelHeader header;

    // 是否有标题
    // private boolean hasTitle = false;

    // 是否有二级标题
    private boolean hasSubTitle = false;

    // 列标题
    private List<ExcelTitle> excelTitle;

    // 是否有表底部备注栏
    private boolean hasFooter = false;

    // 备注
    private ExcelFooter footer;

    // 当单元格为空时填充字符
    private String fillChar = "";

    public boolean isHasSubTitle() {
        return hasSubTitle;
    }

    public void setHasSubTitle(boolean hasSubTitle) {
        this.hasSubTitle = hasSubTitle;
    }

    public String getFileType() {
        return fileType;
    }

    public void setFileType(String fileType) {
        this.fileType = fileType;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public boolean isHasHeader() {
        return hasHeader;
    }

    public void setHasHeader(boolean hasHeader) {
        this.hasHeader = hasHeader;
    }

    public ExcelHeader getHeader() {
        return header;
    }

    public void setHeader(ExcelHeader header) {
        this.header = header;
    }

    public List<ExcelTitle> getExcelTitle() {
        return excelTitle;
    }

    public void setExcelTitle(List<ExcelTitle> excelTitle) {
        this.excelTitle = excelTitle;
    }

    public boolean isHasFooter() {
        return hasFooter;
    }

    public void setHasFooter(boolean hasFooter) {
        this.hasFooter = hasFooter;
    }

    public ExcelFooter getFooter() {
        return footer;
    }

    public void setFooter(ExcelFooter footer) {
        this.footer = footer;
    }

    public String getFillChar() {
        return fillChar;
    }

    public void setFillChar(String fillChar) {
        this.fillChar = fillChar;
    }
}