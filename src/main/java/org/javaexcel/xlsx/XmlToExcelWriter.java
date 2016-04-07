package org.javaexcel.xlsx;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.javaexcel.ExcelWriter;
import org.javaexcel.model.CellMerge;
import org.javaexcel.model.ExcelCellStyle;
import org.javaexcel.model.ExcelFooter;
import org.javaexcel.model.ExcelHeader;
import org.javaexcel.model.ExcelMetaData;
import org.javaexcel.model.ExcelTitle;
import org.javaexcel.util.Const;
import org.javaexcel.util.FileUtils;
import org.javaexcel.util.JsonUtil;
import org.javaexcel.util.UUIDUtil;

/*
 * XLSX文件导出工具类(不支持xls)
 * 先将数据写入临时XML,然后再将XML压缩进EXCEL文件
 * File name   : XmlToExcelWriter.java
 * Description : excelservice-service
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public class XmlToExcelWriter extends ExcelWriter {
    // 存储所有的标题(子级标题)
    private List<ExcelTitle> allTitles = new ArrayList<ExcelTitle>();
    // 存储所有合并单元格
    private List<CellMerge> cellMerges = new ArrayList<CellMerge>();

    private static final double DEFAULTROWHEIGHT = 16;

    private SpreadSheetWriter sw;
    private CellMerge cellMerge;

    /**
     * 导出文件
     * 
     * @throws Exception
     */
    public boolean process(ExcelMetaData metedata, List<Object> datas, String fileName) throws Exception {
        this.metedata = metedata;
        this.allDatas.addAll(datas);
        boolean result = false;

        String tempFile = Files.createTempFile(UUIDUtil.getUUID(), Const.EXCEL_SUFFIX_XLSX).toString();
        String tmpXml = Files.createTempFile(metedata.getSheetName(), Const.XML_SUFFIX).toString();
        if (fileName.endsWith(Const.EXCEL_SUFFIX_XLS)) {
            throw new Exception("Export 2003 version excel is not supported.");
        }

        // 校验fileName的父目录是否存在
        File file = new File(fileName);
        if (Files.notExists(Paths.get(file.getParent()))) {
            throw new Exception("Output directory does not exist.");
        }

        try (OutputStream os = new FileOutputStream(file)) {
            // 建立工作簿和电子表格对象
            wb = new XSSFWorkbook();
            sheet = (XSSFSheet) wb.createSheet(metedata.getSheetName());
            // 持有电子表格数据的xml文件名 例如 /xl/worksheets/sheet1.xml
            String sheetRef = ((XSSFSheet) sheet).getPackagePart().getPartName().getName();

            init();
            initStyle();

            OutputStream out = new FileOutputStream(tempFile);
            wb.write(out);
            wb.close();
            out.close();

            // 生成xml文件
            Writer wr = new FileWriter(tmpXml);
            sw = new SpreadSheetWriter(wr);
            generate();
            wr.close();

            FileUtils.substitute(tempFile, tmpXml, sheetRef.substring(1), os);
            result = true;
        } catch (Exception e) {
            throw e;
        } finally {
            Files.delete(Paths.get(tempFile));
            Files.delete(Paths.get(tmpXml));
        }

        return result;
    }

    private void initStyle() {
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
                cellStyle.setFillForegroundColor(style.getBackgroundColor());
                cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
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

        allTitles.stream().forEach(excelTitle -> {
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
        });
    }

    /**
     * @throws Exception
     * 
     */
    private void init() {
        metedata.getExcelTitle().stream().forEach(et -> {
            if (null != et.getSubTitles() && !et.getSubTitles().isEmpty()) {
                allTitles.addAll(et.getSubTitles());
                columnSize += et.getSubTitles().size();
            } else {
                allTitles.add(et);
                columnSize++;
            }
        });

        // for (int i = 0; i < metedata.getExcelTitle().size(); i++) {
        // ExcelTitle excelTitle = metedata.getExcelTitle().get(i);
        // if (null != excelTitle.getSubTitles() &&
        // !excelTitle.getSubTitles().isEmpty()) {
        // allTitles.addAll(excelTitle.getSubTitles());
        // columnSize += excelTitle.getSubTitles().size();
        // continue;
        // }
        // allTitles.add(excelTitle);
        // columnSize++;
        // }
    }

    /**
     * @throws Exception
     * 
     */
    private void generate() throws Exception {
        // 电子表格开始
        sw.beginSheet();

        // 设置列宽
        sw.beginSetColWidth();
        for (ExcelTitle exct : this.allTitles) {
            if (exct.getColumnWidth() > 0) {
                sw.setColWidthBeforeSheet(exct.getIndex(), exct.getColumnWidth());
                continue;
            }
            sw.setColWidthBeforeSheet(exct.getIndex(), exct.getDisplayName().length() * 3.2);
        }
        sw.endSetColWidth();

        sw.beginSheetData();
        // 写大表头
        writeHeader();

        // 写标题
        writeTitle();

        // 写数据
        writeData();

        // 写备注
        writeFooter();
        sw.endSheetData();

        // 设置合并单元格
        if (!cellMerges.isEmpty()) {
            sw.beginMergerCell(cellMerges.size());
            for (CellMerge mc : cellMerges) {
                sw.setMergeCell(mc.getBeginColumn(), mc.getBeginCell(), mc.getEndColumn(), mc.getEndCell());
            }
            sw.endMergerCell();
        }

        // 电子表格结束
        sw.endSheet();
    }

    /**
     * @throws IOException
     * 
     */
    private void writeFooter() throws IOException {
        if (this.metedata.isHasFooter() && null != this.metedata.getFooter()) {
            cellStyle = this.stylesMap.get("footerStyle");
            sw.insertRowWithHeight(rownum, columnSize, this.metedata.getFooter().getRowHeight());
            for (int i = 0; i < this.columnSize; i++) {
                if (0 == i) {
                    sw.createCell(i, metedata.getFooter().getRemarks(), cellStyle.getIndex());
                    cellMerge = new CellMerge(rownum, i, rownum, columnSize - 1);
                    cellMerges.add(cellMerge);
                    continue;
                }
                sw.createCell(i, cellStyle.getIndex());
            }
            sw.endRow();
        }
    }

    /**
     * @throws IOException
     * 
     */
    private void writeHeader() throws IOException {
        if (this.metedata.isHasHeader() && null != this.metedata.getHeader()) {
            cellStyle = this.stylesMap.get("headerStyle");
            sw.insertRowWithHeight(rownum, columnSize, metedata.getHeader().getRowHeight());
            for (int i = 0; i < columnSize; i++) {
                if (0 == i) {
                    sw.createCell(i, metedata.getHeader().getHeaderName(), cellStyle.getIndex());
                    cellMerge = new CellMerge(rownum, i, rownum, columnSize - 1);
                    cellMerges.add(cellMerge);
                    continue;
                }
                sw.createCell(i, cellStyle.getIndex());
            }

            sw.endRow();
            rownum++;
        }
    }

    /**
     * @throws IOException
     * @throws ParseException
     * 
     */
    @SuppressWarnings("unchecked")
    private void writeData() throws IOException, ParseException {
        if (null == this.allDatas || this.allDatas.isEmpty()) {
            return;
        }

        for (Object object : allDatas) {
            Map<String, Object> dataMap = JsonUtil.stringToBean(JsonUtil.beanToString(object), Map.class);
            if (null == dataMap || dataMap.isEmpty()) {
                continue;
            }

            int rowsize = getColumns(dataMap);
            int maxRow = rownum + rowsize - 1;
            if (rowsize > 0) {
                for (int i = 0; i < rowsize; i++) {
                    sw.insertRowWithHeight(rownum, columnSize, DEFAULTROWHEIGHT);
                    for (ExcelTitle eh : this.metedata.getExcelTitle()) {
                        Object obj = dataMap.get(eh.getName());
                        cellStyle = this.stylesMap.get("cellstyle_" + eh.getIndex());
                        if (eh.isMerge()) {
                            if (0 == i) {
                                cellMerge = new CellMerge(rownum, eh.getIndex(), maxRow, eh.getIndex());
                                cellMerges.add(cellMerge);
                                sw.createCell(eh.getIndex(), obj.toString(), cellStyle.getIndex());
                                continue;
                            }
                            sw.createCell(eh.getIndex(), cellStyle.getIndex());
                        } else if (!eh.getSubTitles().isEmpty() && (obj instanceof List)) {
                            List<Object> list = (List<Object>) obj;
                            Map<String, Object> detailData = (Map<String, Object>) list.get(i);
                            for (ExcelTitle ele : eh.getSubTitles()) {
                                cellStyle = this.stylesMap.get("cellstyle_" + ele.getIndex());
                                sw.createCell(ele.getIndex(), detailData.get(ele.getName()).toString(), cellStyle.getIndex());
                            }
                        }
                    }
                    rownum++;
                    sw.endRow();
                }
            } else {
                sw.insertRowWithHeight(rownum++, columnSize, DEFAULTROWHEIGHT);
                for (ExcelTitle eh : this.allTitles) {
                    cellStyle = this.stylesMap.get("cellstyle_" + eh.getIndex());
                    sw.createCell(eh.getIndex(), dataMap.get(eh.getName()).toString(), cellStyle.getIndex());
                }
                sw.endRow();
            }
        }
    }

    @SuppressWarnings("rawtypes")
    private static int getColumns(Map<String, Object> dataMap) {
        for (Object obj : dataMap.values()) {
            if (obj instanceof List) {
                return ((List) obj).size();
            }
        }
        return 0;
    }

    private void writeTitle() throws IOException {
        if (null == metedata.getExcelTitle() || metedata.getExcelTitle().isEmpty()) {
            return;
        }

        // 写EXCEL表头
        if (metedata.isHasSubTitle()) {
            cellStyle = this.stylesMap.get("titleStyle");
            for (int i = 0; i < 2; i++) {
                sw.insertRowWithHeight(rownum, columnSize, DEFAULTROWHEIGHT);
                for (ExcelTitle excelTitle : metedata.getExcelTitle()) {
                    if (excelTitle.isMerge()) {
                        if (0 == i) {
                            cellMerge = new CellMerge(rownum, excelTitle.getIndex(), rownum + 1, excelTitle.getIndex());
                            cellMerges.add(cellMerge);
                            sw.createCell(excelTitle.getIndex(), excelTitle.getDisplayName(), cellStyle.getIndex());
                            continue;
                        }
                        sw.createCell(excelTitle.getIndex(), cellStyle.getIndex());
                    } else if (null != excelTitle.getSubTitles() && !excelTitle.getSubTitles().isEmpty()) {
                        for (int j = 0; j < excelTitle.getSubTitles().size(); j++) {
                            ExcelTitle ct = excelTitle.getSubTitles().get(j);
                            if (0 == i) {
                                if (0 == j) {
                                    cellMerge = new CellMerge(rownum, ct.getIndex(), rownum,
                                            excelTitle.getSubTitles().get(excelTitle.getSubTitles().size() - 1).getIndex());
                                    cellMerges.add(cellMerge);
                                    sw.createCell(ct.getIndex(), ct.getDisplayName(), cellStyle.getIndex());
                                    continue;
                                }
                                sw.createCell(ct.getIndex(), cellStyle.getIndex());
                            } else {
                                sw.createCell(ct.getIndex(), ct.getDisplayName(), cellStyle.getIndex());
                            }
                        }
                    }
                }
                sw.endRow();
                rownum++;
            }
        } else {
            sw.insertRowWithHeight(rownum++, columnSize, DEFAULTROWHEIGHT);
            for (ExcelTitle et : this.allTitles) {
                sw.createCell(et.getIndex(), et.getDisplayName(), cellStyle.getIndex());
            }
            sw.endRow();
        }
    }
}