package org.javaexcel;

import java.util.List;

import org.javaexcel.model.ExcelMetaData;

/*
 * File name   : ExcelWriter.java
 * @Copyright  : www.quancheng-ec.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月1日
 */
public abstract class ExcelWriter {
    /**
     * 导出数据到EXCEL
     * 
     * @param metadata
     * @param datas
     * @param filePath
     * @return
     */
    public abstract boolean write(ExcelMetaData metadata, List<Object> datas, String filePath);
}