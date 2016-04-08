package org.javaexcel.xls;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.javaexcel.ExcelWriter;
import org.javaexcel.model.ExcelMetaData;
import org.javaexcel.model.ExcelTitle;
import org.javaexcel.util.Const;
import org.javaexcel.util.FileUtils;
import org.javaexcel.util.JsonUtil;

/*
 * File name   : ExcelWriterImpl.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月6日
 */
public class DataToExcelWriter extends ExcelWriter {
    // 存储大标题
    private List<ExcelTitle> bigheaders = new ArrayList<ExcelTitle>();

    // 存储无子标题的Title
    private List<ExcelTitle> mergeTitles = new ArrayList<ExcelTitle>();
    private List<CellRangeAddress> cellRanges = new ArrayList<CellRangeAddress>();

    @Override
    public boolean process(ExcelMetaData metedata, List<Object> datas,
            String fileName) throws Exception {
        this.metedata = metedata;
        this.allDatas.addAll(datas);
        boolean result = false;

        // 校验fileName的父目录是否存在
        if (FileUtils.isExistsOfParentDir(fileName)) {
            throw new Exception("Output directory does not exist.");
        }

        try (FileOutputStream out = new FileOutputStream(fileName)) {
            initWorkbook(fileName);
            sheet = wb.createSheet(this.metedata.getSheetName());

            init();
            initStyle();

            // 写大表头
            writeHeader();

            // 写表头
            writeTitle();

            // 写数据
            writeData();

            writeFooter();

            setAllRangeBorder();

            wb.write(out);
            wb.close();
            result = true;
        } catch (Exception e) {
            throw e;
        }
        return result;
    }

    /**
     * 获取 Workbook
     */
    public void initWorkbook(String fileName) {
        if (fileName.toLowerCase().endsWith(Const.EXCEL_SUFFIX_XLS)) {
            wb = new HSSFWorkbook();
        } else {
            wb = new XSSFWorkbook();
        }
    }

    private void writeHeader() {
        if (this.metedata.isHasHeader() && null != this.metedata.getHeader()) {
            cellStyle = this.getStyle("headerStyle");
            row = sheet.createRow(rownum);
            CellRangeAddress address = new CellRangeAddress(rownum, rownum, 0, columnSize - 1);
            sheet.addMergedRegion(address);
            cellRanges.add(address);
            row.setHeightInPoints(metedata.getHeader().getRowHeight());

            CellUtil.createCell(row, rownum, metedata.getHeader().getHeaderName(), cellStyle);
            rownum++;
        }
    }

    private void writeFooter() {
        if (this.metedata.isHasFooter() && null != this.metedata.getFooter()) {
            cellStyle = this.getStyle("footerStyle");
            row = sheet.createRow(rownum);
            CellRangeAddress address = new CellRangeAddress(rownum, rownum, 0, columnSize - 1);
            sheet.addMergedRegion(address);
            cellRanges.add(address);
            row.setHeightInPoints(metedata.getFooter().getRowHeight());
            CellUtil.createCell(row, 0, metedata.getFooter().getRemarks(), cellStyle);
        }
    }

    /**
     * 
     */
    private void setAllRangeBorder() {
        for (CellRangeAddress address : this.cellRanges) {
            this.setBorder(address);
        }
    }

    private void writeTitle() {
        if (null == metedata.getExcelTitle() || metedata.getExcelTitle().isEmpty()) {
            return;
        }

        cellStyle = this.getStyle("titleStyle");
        if (metedata.isHasSubTitle() && null != this.metedata.getExcelTitle() && !this.metedata.getExcelTitle().isEmpty()) {
            List<ExcelTitle> titles = metedata.getExcelTitle();
            CellRangeAddress address = null;
            for (ExcelTitle ex : titles) {
                if (null != ex.getSubTitles() && ex.getSubTitles().size() >= 2) {
                    address = new CellRangeAddress(rownum, rownum, ex.getSubTitles().get(0).getIndex(), ex.getSubTitles().get(ex.getSubTitles().size() - 1).getIndex());

                    sheet.addMergedRegion(address);
                    cellRanges.add(address);
                    bigheaders.add(ex);
                    continue;
                }

                // 合并头两行的某一列
                address = new CellRangeAddress(rownum, rownum + 1, ex.getIndex(), ex.getIndex());
                sheet.addMergedRegion(address);
                cellRanges.add(address);
                mergeTitles.add(ex);
            }

            // 写Excel表头
            Row oneRow = sheet.createRow(rownum++);
            oneRow.setHeightInPoints(DEFAULTROWHEIGHT);
            Row twoRow = sheet.createRow(rownum++);
            twoRow.setHeightInPoints(DEFAULTROWHEIGHT);
            for (ExcelTitle ete : titles) {
                if (null != ete.getSubTitles() && ete.getSubTitles().size() >= 2) {
                    CellUtil.createCell(oneRow, ete.getSubTitles().get(0).getIndex(), ete.getDisplayName(), cellStyle);
                    for (ExcelTitle subTitle : ete.getSubTitles()) {
                        createCell(twoRow, subTitle);
                    }
                } else {
                    createCell(oneRow, ete);
                }
            }
        } else {
            row = sheet.createRow(rownum++);
            row.setHeightInPoints(DEFAULTROWHEIGHT);
            for (ExcelTitle eh : this.metedata.getExcelTitle()) {
                CellUtil.createCell(row, eh.getIndex(), eh.getName(), cellStyle);
            }
        }
    }

    private void createCell(Row row, ExcelTitle ete) {
        if (ete.getColumnWidth() > 0) {
            sheet.setColumnWidth(ete.getIndex(), ete.getColumnWidth() * 256);
        } else {
            sheet.setColumnWidth(ete.getIndex(), ete.getDisplayName().length() * 4 * 256);
        }
        CellUtil.createCell(row, ete.getIndex(), ete.getDisplayName(), cellStyle);
    }

    @SuppressWarnings({ "unchecked", "rawtypes" })
    private void writeData() {
        for (Object object : allDatas) {
            Map<String, Object> dataMap = JsonUtil.stringToBean(JsonUtil.beanToString(object), Map.class);
            if (null == dataMap || dataMap.isEmpty()) {
                continue;
            }

            int rowsize = getColumns(dataMap);
            int maxRow = rownum + rowsize - 1;
            if (rowsize > 0) {
                // 需要处理行的数据合并
                for (ExcelTitle ex : mergeTitles) {
                    CellRangeAddress address = new CellRangeAddress(rownum, maxRow, ex.getIndex(), ex.getIndex());
                    cellRanges.add(address);
                    sheet.addMergedRegion(address);
                }

                row = sheet.createRow(rownum++);
                row.setHeightInPoints(DEFAULTROWHEIGHT);
                for (ExcelTitle eh : mergeTitles) {
                    createCell(eh, dataMap.get(eh.getName()));
                }

                for (int i = 0; i < rowsize; i++) {
                    for (ExcelTitle ele : bigheaders) {
                        Object obj = dataMap.get(ele.getName());
                        if (obj instanceof List) {
                            Map<String, Object> detailData = (Map<String, Object>) ((List) obj).get(i);
                            ele.getSubTitles().stream().forEach(exte -> {
                                createCell(exte, detailData.get(exte.getName()));
                            });
                        }
                    }

                    // 最后一次循环需要处理迭代索引
                    if (i != rowsize - 1) {
                        row = sheet.createRow(rownum++);
                        row.setHeightInPoints(DEFAULTROWHEIGHT);
                    }
                }
            } else {
                row = sheet.createRow(rownum++);
                row.setHeightInPoints(DEFAULTROWHEIGHT);
                createAllCell(row, dataMap);
            }
        }
    }

    private void createCell(ExcelTitle ete, Object obj) {
        cellStyle = this.getStyle("cellstyle_" + ete.getIndex());
        String result = "";
        if (JsonUtil.isEmpty(obj)) {
            if (!JsonUtil.isEmpty(ete.getFillChar())) {
                result = ete.getFillChar();
            }

            CellUtil.createCell(row, ete.getIndex(), result, cellStyle);
            return;
        }

        cell = row.createCell(ete.getIndex());
        cell.setCellStyle(cellStyle);
        switch (ete.getDataType()) {
            case NUMBER:
            case INT:
            case MONEY:
            case PERCENT:
                cell.setCellValue(Double.valueOf(obj.toString()));
                break;
            default:
                result = obj.toString();
                cell.setCellValue(result);
                break;
        }
    }

    private void createAllCell(Row row, Map<String, Object> data) {
        allTitles.stream().forEach(eh -> {
            cellStyle = this.getStyle("cellstyle_" + eh.getIndex());
            CellUtil.createCell(row, eh.getIndex(), data.get(eh.getName()) + "", cellStyle);
        });
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

    public void setBorder(CellRangeAddress cellRangeAddress) {
        RegionUtil.setBorderLeft(CellStyle.BORDER_THIN, cellRangeAddress, sheet, wb);
        RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, cellRangeAddress, sheet, wb);
        RegionUtil.setBorderRight(CellStyle.BORDER_THIN, cellRangeAddress, sheet, wb);
        RegionUtil.setBorderTop(CellStyle.BORDER_THIN, cellRangeAddress, sheet, wb);
    }
}