package org.javaexcel.xlsx;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.javaexcel.ExcelWriter;
import org.javaexcel.ExcelWriterFactory;
import org.javaexcel.model.CellType;
import org.javaexcel.model.ExcelCellStyle;
import org.javaexcel.model.ExcelColor;
import org.javaexcel.model.ExcelHeader;
import org.javaexcel.model.ExcelMetaData;
import org.javaexcel.model.ExcelTitle;
import org.javaexcel.model.ExcelType;
import org.junit.Before;
import org.junit.Test;

/*
 * File name   : SimpleReportExportTest.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月6日
 */
public class SimpleReportExportTest {
    private static final int ROWS = 10;
    private ExcelMetaData metadata;
    private List<Object> datas;

    @Test
    public void test() {
        long begTime = System.currentTimeMillis();
        try {
            ExcelWriter writer = ExcelWriterFactory.getWriter(ExcelType.XLS);
            writer.process(metadata, datas, "/Users/Robert/Desktop/QA_test/user.xls");
        } catch (Exception e) {
            e.printStackTrace();
        }

        System.out.println("Total:" + (System.currentTimeMillis() - begTime) + "ms");
    }

    @Before
    public void setUp() throws Exception {
        metadata = new ExcelMetaData();
        metadata.setFileName("SimpleReport");
        metadata.setSheetName("SimpleReport");

        // 设置大表头
        this.metadata.setHasHeader(true);
        ExcelHeader header = new ExcelHeader();
        header.setHeaderName("报表");
        ExcelCellStyle hs = new ExcelCellStyle();
        hs.setAlign(XSSFCellStyle.ALIGN_CENTER);
        hs.setVerticalAlign(XSSFCellStyle.ALIGN_CENTER);
        hs.setSize((short) 42);
        hs.setColor(ExcelColor.BLACK);
        header.setCellStyle(hs);
        this.metadata.setHeader(header);

        // 初始化数据
        structureMetaData();
        constructdata();
    }

    /**
     * 
     */
    private void constructdata() {
        this.datas = new ArrayList<Object>();
        User user = null;
        for (int i = 0; i < ROWS; i++) {
            user = new User(i + 1, "张晓明" + i, new Date(), "18913538888", "上海市浦东新区康桥御桥路200号", 0.45);
            this.datas.add(user);
        }
    }

    /**
     * 
     */
    private void structureMetaData() {
        // 设置表头
        List<ExcelTitle> titles = new ArrayList<ExcelTitle>();
        int rownum = 0;

        ExcelTitle t1 = new ExcelTitle();
        t1.setIndex(rownum++);
        t1.setName("id");
        t1.setDisplayName("ID");
        t1.setDataType(CellType.INT);
        t1.setColumnWidth(8);
        titles.add(t1);

        ExcelTitle t2 = new ExcelTitle();
        t2.setIndex(rownum++);
        t2.setName("name");
        t2.setDisplayName("姓名");
        t2.setDataType(CellType.TEXT);
        t2.setColumnWidth(10);
        titles.add(t2);

        ExcelTitle t3 = new ExcelTitle();
        t3.setIndex(rownum++);
        t3.setName("birthday");
        t3.setDisplayName("出生日期");
        t3.setColumnWidth(20);
        t3.setDataType(CellType.DATE);
        titles.add(t3);

        ExcelTitle t4 = new ExcelTitle();
        t4.setIndex(rownum++);
        t4.setName("telphone");
        t4.setDisplayName("电话");
        t4.setDataType(CellType.TEXT);
        t4.setColumnWidth(15);
        titles.add(t4);

        ExcelTitle t5 = new ExcelTitle();
        t5.setIndex(rownum++);
        t5.setName("address");
        t5.setDisplayName("联系地址");
        t5.setDataType(CellType.TEXT);
        t5.setColumnWidth(55);
        titles.add(t5);

        ExcelTitle t6 = new ExcelTitle();
        t6.setIndex(rownum++);
        t6.setName("percent");
        t6.setDisplayName("联系地址");
        t6.setDataType(CellType.PERCENT);
        t6.setColumnWidth(55);
        titles.add(t6);

        this.metadata.setExcelTitle(titles);
        this.metadata.setHasSubTitle(false);
    }

    class User {
        private Integer id;
        private String name;
        private Date birthday;
        private String telphone;
        private String address;
        private double percent;

        public double getPercent() {
            return percent;
        }

        public void setPercent(double percent) {
            this.percent = percent;
        }

        /**
         * @param id
         * @param name
         * @param birthday
         * @param telphone
         * @param address
         */
        public User(Integer id, String name, Date birthday, String telphone, String address, double percent) {
            super();
            this.id = id;
            this.name = name;
            this.birthday = birthday;
            this.telphone = telphone;
            this.address = address;
            this.percent = percent;
        }

        public Integer getId() {
            return id;
        }

        public void setId(Integer id) {
            this.id = id;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Date getBirthday() {
            return birthday;
        }

        public void setBirthday(Date birthday) {
            this.birthday = birthday;
        }

        public String getTelphone() {
            return telphone;
        }

        public void setTelphone(String telphone) {
            this.telphone = telphone;
        }

        public String getAddress() {
            return address;
        }

        public void setAddress(String address) {
            this.address = address;
        }
    }
}