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
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.javaexcel.model.CellMerge;
import org.javaexcel.model.ExcelMetaData;
import org.javaexcel.model.ExcelTitle;
import org.javaexcel.util.Const;
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
    private SpreadSheetWriter sw;
    private List<Map<String, Object>> allDatas;
    private ExcelMetaData metedata;
    private CellMerge cellMerge;
    private List<CellMerge> cellMerges = new ArrayList<CellMerge>();
    private int rownum = 0;
    private int columnSize = 0;

    /**
     * 导出文件
     * 
     * @throws Exception
     */
    public void process(ExcelMetaData metedata, List<Map<String, Object>> datas, String fileName) throws Exception {
        this.metedata = metedata;
        this.allDatas = datas;

        init();

        String tempFile = Files.createTempFile(UUIDUtil.getUUID(), Const.EXCEL_SUFFIX_XLSX).toString();
        String tmpXml = Files.createTempFile(metedata.getSheetName(), Const.XML_SUFFIX).toString();
        try (OutputStream os = new FileOutputStream(fileName)) {
            // 建立工作簿和电子表格对象
            Workbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet) wb.createSheet(this.metedata.getSheetName());
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
    private void init() {
        for (ExcelTitle excelTitle : metedata.getExcelTitle()) {
            if (excelTitle.isHasSubTitle() && null != excelTitle.getSubTitles()) {
                columnSize += excelTitle.getSubTitles().size();
                continue;
            }
            columnSize++;
        }
    }

    /**
     * @throws Exception
     * 
     */
    private void generate() throws Exception {
        // 电子表格开始
        sw.beginSheet();
        sw.beginSheetData();

        // 写大表头
        writeBigTitle();

        // 写标题
        writeHeader();

        // 写数据
        writeData();

        // 写备注

        sw.endSheetData();

        if (null != cellMerges && !cellMerges.isEmpty()) {
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
    private void writeBigTitle() throws IOException {
        if (this.metedata.isHasHeader()) {
            sw.insertRowWithHeight(rownum, columnSize - 1, 45);
            sw.createCell(rownum, metedata.getHeader().getHeaderName());
            sw.endRow();

            cellMerge = new CellMerge(rownum, rownum, rownum, columnSize - 1);
            cellMerges.add(cellMerge);

            rownum++;
        }
    }

    /**
     * @throws IOException
     * 
     */
    @SuppressWarnings("unchecked")
    private void writeData() throws IOException {
        if (null == this.allDatas || this.allDatas.isEmpty()) {
            return;
        }

        for (Map<String, Object> data : allDatas) {
            int rowsize = getColumns(data);
            int maxRow = rownum + rowsize - 1;
            if (rowsize > 0) {
                for (int i = 0; i < rowsize; i++) {
                    // sw.insertRow(rownum);
                    sw.insertRowWithHeight(rownum, columnSize, DEFAULTROWHEIGHT);
                    for (ExcelTitle eh : this.metedata.getExcelTitle()) {
                        Object obj = data.get(eh.getName());
                        if (eh.isMerge()) {
                            if (0 == i) {
                                cellMerge = new CellMerge(rownum, eh.getIndex(), maxRow, eh.getIndex());
                                cellMerges.add(cellMerge);

                                sw.createCell(eh.getIndex(), obj.toString());
                            }

                            continue;
                        } else if (!eh.getSubTitles().isEmpty() && (obj instanceof List)) {
                            List<Object> list = (List<Object>) obj;
                            Map<String, Object> detailData = (Map<String, Object>) list.get(i);
                            for (ExcelTitle ele : eh.getSubTitles()) {
                                sw.createCell(ele.getIndex(), detailData.get(ele.getName()).toString());
                            }
                        }
                    }
                    rownum++;
                    sw.endRow();
                }
            } else {
                // sw.insertRow(rownum++);
                sw.insertRowWithHeight(rownum++, columnSize, DEFAULTROWHEIGHT);
                for (ExcelTitle eh : this.metedata.getExcelTitle()) {
                    sw.createCell(eh.getIndex(), data.get(eh.getName()).toString());
                }
                sw.endRow();
            }
        }
    }

    @SuppressWarnings("rawtypes")
    private static int getColumns(Map<String, Object> data) {
        for (Object obj : data.values()) {
            if (obj instanceof List) {
                return ((List) obj).size();
            }
        }
        return 0;
    }

    private void writeHeader() throws IOException {
        if (null == metedata.getExcelTitle() || metedata.getExcelTitle().isEmpty()) {
            return;
        }

        // 写EXCEL表头
        if (metedata.isHasSubTitle()) {
            for (int i = 0; i < 2; i++) {
                // sw.insertRow(rownum);
                sw.insertRowWithHeight(rownum, columnSize, DEFAULTROWHEIGHT);
                for (ExcelTitle excelTitle : metedata.getExcelTitle()) {
                    if (excelTitle.isMerge()) {
                        if (0 == i) {
                            cellMerge = new CellMerge(rownum, excelTitle.getIndex(), rownum + 1, excelTitle.getIndex());
                            cellMerges.add(cellMerge);
                            sw.createCell(excelTitle.getIndex(), excelTitle.getDisplayName());
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
                                    sw.createCell(ct.getIndex(), ct.getDisplayName());
                                }
                                continue;
                            }

                            sw.createCell(ct.getIndex(), ct.getDisplayName());
                        }
                    }
                }
                sw.endRow();
                rownum++;
            }
        } else {
            sw.insertRowWithHeight(rownum++, columnSize, DEFAULTROWHEIGHT);
            for (ExcelTitle et : metedata.getExcelTitle()) {
                sw.createCell(et.getIndex(), et.getDisplayName());
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