package org.javaexcel.xlsx;

import java.io.IOException;
import java.io.Writer;

import org.apache.poi.hssf.util.CellReference;
import org.javaexcel.model.CellMerge;

/*
 * File name   : SpreadSheetWriter.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月2日
 */
public class SpreadSheetWriter {
    // XML Encode
    private static final String[] xmlCode = new String[256];
    private Writer _out;
    private int _rownum;
    private static String LINE_SEPARATOR = System.getProperty("line.separator");

    static {
        // Special characters
        xmlCode['\''] = "'";
        // double quote
        xmlCode['\"'] = "\"";
        // ampersand
        xmlCode['&'] = "&";
        // lower than
        xmlCode['<'] = "<";
        // greater than
        xmlCode['>'] = ">";
    }

    public SpreadSheetWriter(Writer out) {
        _out = out;
    }

    public void beginSheet() throws IOException {
        _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        _out.write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + LINE_SEPARATOR);
    }

    public void endSheet() throws IOException {
        _out.write("</worksheet>");
    }

    public void beginSheetData() throws IOException {
        _out.write("<sheetData>" + LINE_SEPARATOR);
    }

    public void endSheetData() throws IOException {
        _out.write("</sheetData>");
    }

    /**
     * 插入新行
     * 
     * @param rownum(以0开始)
     */
    public void insertRow(int rownum) throws IOException {
        _out.write("<row r=\"" + (rownum + 1) + "\">" + LINE_SEPARATOR);
        this._rownum = rownum;
    }

    /**
     * 插入行且设置高度
     * 
     * @param rownum(以0开始)
     * @param columnNum
     * @param height
     * @throws IOException
     */
    public void insertRowWithHeight(int rownum, int columnNum, double height) throws IOException {
        this._out.write("<row r=\"" + (rownum + 1) + "\" spans=\"1:" + columnNum + "\" ht=\"" + height
                + "\" customHeight=\"1\">" + LINE_SEPARATOR);
        this._rownum = rownum;
    }

    /**
     * 插入行结束标志
     */
    public void endRow() throws IOException {
        _out.write("</row>" + LINE_SEPARATOR);
    }

    /**
     * 开始设置列宽
     * 
     * @throws IOException
     */
    public void beginSetColWidth() throws IOException {
        this._out.write("<cols>" + LINE_SEPARATOR);
    }

    /**
     * 设置列宽下标从0开始
     * 
     * @param columnIndex
     * @param columnWidth
     * @throws IOException
     */
    public void setColWidthBeforeSheet(int columnIndex, double columnWidth) throws IOException {
        this._out.write("<col min=\"" + (columnIndex + 1) + "\" max=\""
                + (columnIndex + 1) + "\" width=\"" + columnWidth
                + "\" customWidth=\"1\"/>" + LINE_SEPARATOR);
    }

    /**
     * 设置列宽结束
     * 
     * @throws IOException
     */
    public void endSetColWidth() throws IOException {
        this._out.write("</cols>" + LINE_SEPARATOR);
    }

    /**
     * 合并单元格开始标记
     * 
     * @throws IOException
     */
    public void beginMergerCell(int count) throws IOException {
        this._out.write("<mergeCells  count=\"" + count + "\">" + LINE_SEPARATOR);
    }

    /**
     * 合并单元格(下标从0开始)
     * 
     * @param beginColumn
     * @param beginCell
     * @param endColumn
     * @param endCell
     * @throws IOException
     */
    public void setMergeCell(int beginColumn, int beginCell, int endColumn,
            int endCell) throws IOException {
        this._out.write("<mergeCell ref=\"" + CellReference.convertNumToColString(beginCell)
                + (beginColumn + 1) + ":" + CellReference.convertNumToColString(endCell)
                + (endColumn + 1) + "\"/>" + LINE_SEPARATOR);
    }

    public void setMergeCell(CellMerge mergeCell) throws IOException {
        _out.write(mergeCell.toString());
    }

    /**
     * 合并单元格结束标记
     * 
     * @throws IOException
     */
    public void endMergerCell() throws IOException {
        this._out.write("</mergeCells>" + LINE_SEPARATOR);
    }

    /**
     * 插入新列
     * 
     * @param columnIndex
     * @param value
     * @param styleIndex
     * @throws IOException
     */
    public void createCell(int columnIndex, String value, int styleIndex)
            throws IOException {
        String ref = new CellReference(_rownum, columnIndex).formatAsString();
        _out.write("<c r=\"" + ref + "\" t=\"inlineStr\"");
        if (styleIndex != -1) {
            _out.write(" s=\"" + styleIndex + "\"");
        }
        _out.write(">");
        _out.write("<is><t>" + encoderXML(value) + "</t></is>");
        _out.write("</c>");
    }

    /**
     * 插入一个Cell(包含值)
     * 
     * @param columnIndex
     * @param value
     * @throws IOException
     */
    public void createCell(int columnIndex, String value)
            throws IOException {
        createCell(columnIndex, value, -1);
    }

    /**
     * 插入一个Cell(不包含值,合并单元格)
     * 
     * @param columnIndex
     * @throws IOException
     */
    public void createCell(int columnIndex, int styleIndex) throws IOException {
        String ref = new CellReference(_rownum, columnIndex).formatAsString();
        _out.write("<c r=\"" + ref + "\" s=\"" + styleIndex + "\" />");
    }

    public void createCell(int columnIndex, double value, int styleIndex)
            throws IOException {
        String ref = new CellReference(_rownum, columnIndex).formatAsString();
        _out.write("<c r=\"" + ref + "\" t=\"n\"");
        if (styleIndex != -1) {
            _out.write(" s=\"" + styleIndex + "\"");
        }
        _out.write(">");
        _out.write("<v>" + value + "</v>");
        _out.write("</c>");
    }

    /**
     * <p>
     * Encode the given text into xml.
     * </p>
     * 
     * @param string
     *            the text to encode
     * @return the encoded string
     */
    public static String encoderXML(String string) {
        if (string == null)
            return "";
        int n = string.length();
        char character;
        String xmlchar;
        StringBuffer buffer = new StringBuffer();
        // loop over all the characters of the String.
        for (int i = 0; i < n; i++) {
            character = string.charAt(i);
            // the xmlcode of these characters are added to a StringBuffer
            // one by one
            try {
                xmlchar = xmlCode[character];
                if (xmlchar == null) {
                    buffer.append(character);
                } else {
                    buffer.append(xmlCode[character]);
                }
            } catch (ArrayIndexOutOfBoundsException aioobe) {
                buffer.append(character);
            }
        }
        return buffer.toString();
    }
}