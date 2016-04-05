package org.javaexcel.xlsx;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.javaexcel.model.CellMerge;
import org.javaexcel.model.ExcelCellStyle;
import org.javaexcel.model.ExcelFooter;
import org.javaexcel.model.ExcelHeader;
import org.javaexcel.model.ExcelMetaData;
import org.javaexcel.model.ExcelTitle;
import org.javaexcel.util.Const;
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
public class XmlToExcelWriter {
    private static final double DEFAULTROWHEIGHT = 16;
    private Workbook wb;
    private XSSFSheet sheet;
    private SpreadSheetWriter sw;
    private List<Object> allDatas;
    private ExcelMetaData metedata;
    private CellMerge cellMerge;
    private List<CellMerge> cellMerges = new ArrayList<CellMerge>();
    private Map<String, XSSFCellStyle> stylesMap = new HashMap<String, XSSFCellStyle>();
    private int rownum = 0;
    private int columnSize = 0;
    private XSSFCellStyle xssfCellStyle;

    // 存储所有的标题(子级标题)
    private List<ExcelTitle> allTitles = new ArrayList<ExcelTitle>();

    /**
     * 导出文件
     * 
     * @throws Exception
     */
    public void process(ExcelMetaData metedata, List<Object> datas, String fileName) throws Exception {
        this.metedata = metedata;
        this.allDatas = datas;

        String tempFile = Files.createTempFile(UUIDUtil.getUUID(), Const.EXCEL_SUFFIX_XLSX).toString();
        String tmpXml = Files.createTempFile(metedata.getSheetName(), Const.XML_SUFFIX).toString();

        // 校验fileName的父目录是否存在

        try (OutputStream os = new FileOutputStream(fileName)) {
            // 建立工作簿和电子表格对象
            wb = new XSSFWorkbook();
            sheet = (XSSFSheet) wb.createSheet(metedata.getSheetName());

            init();
            initStyle();

            // 持有电子表格数据的xml文件名 例如 /xl/worksheets/sheet1.xml
            String sheetRef = sheet.getPackagePart().getPartName().getName();

            OutputStream out = new FileOutputStream(tempFile);
            wb.write(out);
            wb.close();
            out.close();

            // 生成xml文件
            Writer wr = new FileWriter(tmpXml);
            sw = new SpreadSheetWriter(wr);
            generate();
            wr.close();

            substitute(tempFile, tmpXml, sheetRef.substring(1), os);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            Files.delete(Paths.get(tempFile));
            Files.delete(Paths.get(tmpXml));
        }
    }

    /**
     * 
     */
    private void initStyle() {
        ExcelHeader header = metedata.getHeader();
        if (metedata.isHasHeader() && null != header) {
            xssfCellStyle = (XSSFCellStyle) wb.createCellStyle();
            ExcelCellStyle hsy = header.getCellStyle();
            if (null != hsy) {
                xssfCellStyle.setAlignment(hsy.getAlign());
                xssfCellStyle.setVerticalAlignment(hsy.getVerticalAlign());
                // xssfCellStyle.setBorderRight(CellStyle.BORDER_THIN);
                // xssfCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
                // xssfCellStyle.setBorderLeft(CellStyle.BORDER_THIN);
                // xssfCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                // xssfCellStyle.setBorderTop(CellStyle.BORDER_THIN);
                // xssfCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
                // xssfCellStyle.setBorderBottom(CellStyle.BORDER_THIN);
                // xssfCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

                XSSFFont font = (XSSFFont) wb.createFont();
                font.setFontHeightInPoints((short) hsy.getSize());
                font.setColor((short) hsy.getColor());
                xssfCellStyle.setFont(font);
            }
            this.stylesMap.put("headerStyle", xssfCellStyle);
        }

        ExcelFooter footer = metedata.getFooter();
        if (metedata.isHasFooter() && null != footer) {
            xssfCellStyle = (XSSFCellStyle) wb.createCellStyle();
            ExcelCellStyle fstyle = footer.getCellStyle();
            if (null != fstyle) {
                xssfCellStyle.setAlignment(fstyle.getAlign());
                xssfCellStyle.setVerticalAlignment(fstyle.getVerticalAlign());

                XSSFFont fot = (XSSFFont) wb.createFont();
                fot.setFontHeightInPoints((short) fstyle.getSize());
                fot.setColor((short) fstyle.getColor());
                xssfCellStyle.setFont(fot);
            }
            this.stylesMap.put("footerStyle", xssfCellStyle);
        }

        xssfCellStyle = (XSSFCellStyle) wb.createCellStyle();
        xssfCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        xssfCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        XSSFFont ff = (XSSFFont) wb.createFont();
        ff.setFontHeightInPoints((short) 12);
        xssfCellStyle.setFont(ff);
        stylesMap.put("titleStyle", xssfCellStyle);

        for (ExcelTitle excelTitle : allTitles) {
            xssfCellStyle = (XSSFCellStyle) wb.createCellStyle();
            xssfCellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);

            ExcelCellStyle style = excelTitle.getCellStyle();
            if (null != style) {
                xssfCellStyle.setAlignment(style.getAlign());
                xssfCellStyle.setVerticalAlignment(style.getVerticalAlign());

                XSSFFont ft = (XSSFFont) wb.createFont();
                ft.setFontHeightInPoints(style.getSize());
                ft.setColor(style.getColor());
                xssfCellStyle.setFont(ft);
            }
            stylesMap.put("cellstyle_" + excelTitle.getIndex(), xssfCellStyle);
        }
    }

    /**
     * 
     */
    private void init() {
        for (ExcelTitle excelTitle : metedata.getExcelTitle()) {
            if (excelTitle.isHasSubTitle() && null != excelTitle.getSubTitles()
                    && !excelTitle.getSubTitles().isEmpty()) {
                allTitles.addAll(excelTitle.getSubTitles());
                columnSize += excelTitle.getSubTitles().size();
                continue;
            }
            allTitles.add(excelTitle);
            columnSize++;
        }

        // this.fmt = (XSSFDataFormat) wb.createDataFormat();
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
            sw.setColWidthBeforeSheet(exct.getIndex(), exct.getDisplayName().length() * 3.5);
        }
        sw.endSetColWidth();

        sw.beginSheetData();
        // 写大表头
        writeBigTitle();

        // 写标题
        writeTitle();

        // 写数据
        writeData();

        // 写备注
        writeFooter();
        sw.endSheetData();

        if (null != cellMerges && !cellMerges.isEmpty()) {
            sw.beginMergerCell(cellMerges.size());
            for (CellMerge mc : cellMerges) {
                sw.setMergeCell(mc.getBeginColumn(), mc.getBeginCell(), mc.getEndColumn(), mc.getEndCell());

                // CellRangeAddress address = new
                // CellRangeAddress(mc.getBeginColumn(), mc.getEndColumn(),
                // mc.getBeginCell(), mc.getEndCell());
                // RegionUtil.setBorderLeft(CellStyle.BORDER_MEDIUM_DASHED,
                // address, this.sheet, this.wb);
                // RegionUtil.setBorderRight(CellStyle.BORDER_MEDIUM_DASHED,
                // address, this.sheet, this.wb);
                // RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED,
                // address, this.sheet, this.wb);
                // RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED,
                // address, this.sheet, this.wb);
                //
                // RegionUtil.setLeftBorderColor(ExcelColor.RED, address, sheet,
                // wb);
                // RegionUtil.setRightBorderColor(ExcelColor.RED, address,
                // sheet, wb);
                // RegionUtil.setTopBorderColor(ExcelColor.RED, address, sheet,
                // wb);
                // RegionUtil.setBottomBorderColor(ExcelColor.RED, address,
                // sheet, wb);
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
        if (this.metedata.isHasFooter()) {
            xssfCellStyle = this.stylesMap.get("footerStyle");

            sw.insertRowWithHeight(rownum, columnSize - 1, 25);
            sw.createCell(0, metedata.getFooter().getRemarks(), xssfCellStyle.getIndex());
            sw.endRow();

            cellMerge = new CellMerge(rownum, 0, rownum, columnSize - 1);
            cellMerges.add(cellMerge);
        }
    }

    /**
     * @throws IOException
     * 
     */
    private void writeBigTitle() throws IOException {
        if (this.metedata.isHasHeader()) {
            xssfCellStyle = this.stylesMap.get("headerStyle");

            sw.insertRowWithHeight(rownum, columnSize - 1, 45);
            sw.createCell(rownum, metedata.getHeader().getHeaderName(), xssfCellStyle.getIndex());
            sw.endRow();

            cellMerge = new CellMerge(rownum, rownum, rownum, columnSize - 1);
            cellMerges.add(cellMerge);

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
                        xssfCellStyle = this.stylesMap.get("cellstyle_" + eh.getIndex());
                        if (eh.isMerge()) {
                            if (0 == i) {
                                cellMerge = new CellMerge(rownum, eh.getIndex(), maxRow, eh.getIndex());
                                cellMerges.add(cellMerge);
                                sw.createCell(eh.getIndex(), obj.toString(), xssfCellStyle.getIndex());
                            }
                            continue;
                        } else if (!eh.getSubTitles().isEmpty() && (obj instanceof List)) {
                            List<Object> list = (List<Object>) obj;
                            Map<String, Object> detailData = (Map<String, Object>) list.get(i);
                            for (ExcelTitle ele : eh.getSubTitles()) {
                                xssfCellStyle = this.stylesMap.get("cellstyle_" + ele.getIndex());
                                sw.createCell(ele.getIndex(), detailData.get(ele.getName()).toString(), xssfCellStyle.getIndex());
                            }
                        }
                    }
                    rownum++;
                    sw.endRow();
                }
            } else {
                sw.insertRowWithHeight(rownum++, columnSize, DEFAULTROWHEIGHT);
                for (ExcelTitle eh : this.allTitles) {
                    xssfCellStyle = this.stylesMap.get("cellstyle_" + eh.getIndex());
                    sw.createCell(eh.getIndex(), dataMap.get(eh.getName()).toString(), xssfCellStyle.getIndex());
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

        xssfCellStyle = this.stylesMap.get("titleStyle");

        // 写EXCEL表头
        if (metedata.isHasSubTitle()) {
            for (int i = 0; i < 2; i++) {
                sw.insertRowWithHeight(rownum, columnSize, DEFAULTROWHEIGHT);
                for (ExcelTitle excelTitle : metedata.getExcelTitle()) {
                    if (excelTitle.isMerge()) {
                        if (0 == i) {
                            cellMerge = new CellMerge(rownum, excelTitle.getIndex(), rownum + 1, excelTitle.getIndex());
                            cellMerges.add(cellMerge);
                            sw.createCell(excelTitle.getIndex(), excelTitle.getDisplayName(), xssfCellStyle.getIndex());
                        }
                        continue;
                    } else if (null != excelTitle.getSubTitles() && !excelTitle.getSubTitles().isEmpty()) {
                        for (int j = 0; j < excelTitle.getSubTitles().size(); j++) {
                            ExcelTitle ct = excelTitle.getSubTitles().get(j);
                            if (0 == i) {
                                if (0 == j) {
                                    cellMerge = new CellMerge(rownum, ct.getIndex(), rownum,
                                            excelTitle.getSubTitles().get(excelTitle.getSubTitles().size() - 1).getIndex());
                                    cellMerges.add(cellMerge);
                                    sw.createCell(ct.getIndex(), ct.getDisplayName(), xssfCellStyle.getIndex());
                                }
                                continue;
                            }

                            sw.createCell(ct.getIndex(), ct.getDisplayName(), xssfCellStyle.getIndex());
                        }
                    }
                }
                sw.endRow();
                rownum++;
            }
        } else {
            sw.insertRowWithHeight(rownum++, columnSize, DEFAULTROWHEIGHT);
            for (ExcelTitle et : this.allTitles) {
                sw.createCell(et.getIndex(), et.getDisplayName(), xssfCellStyle.getIndex());
            }
            sw.endRow();
        }
    }

    @SuppressWarnings("unchecked")
    private static void substitute(String zipfile, String tmpfile,
            String entry, OutputStream out) throws IOException {
        ZipFile zip = new ZipFile(zipfile);
        ZipOutputStream zos = new ZipOutputStream(out);
        Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) zip.entries();
        while (en.hasMoreElements()) {
            ZipEntry ze = en.nextElement();
            if (!ze.getName().equals(entry)) {
                zos.putNextEntry(new ZipEntry(ze.getName()));
                InputStream is = zip.getInputStream(ze);
                copyStream(is, zos);
                is.close();
            }
        }
        zos.putNextEntry(new ZipEntry(entry));
        InputStream is = new FileInputStream(tmpfile);
        copyStream(is, zos);
        zip.close();
        is.close();
        zos.close();
    }

    private static void copyStream(InputStream in, OutputStream out)
            throws IOException {
        byte[] chunk = new byte[1024];
        int count;
        while ((count = in.read(chunk)) >= 0) {
            out.write(chunk, 0, count);
        }
    }
}