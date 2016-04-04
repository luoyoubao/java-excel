package org.javaexcel.xlsx;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.javaexcel.model.CellType;
import org.javaexcel.model.ExcelFooter;
import org.javaexcel.model.ExcelHeader;
import org.javaexcel.model.ExcelMetaData;
import org.javaexcel.model.ExcelTitle;
import org.javaexcel.util.JsonUtil;
import org.junit.Before;
import org.junit.Test;

/*
 * File name   : XmlToExcelWriterTest.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月2日
 */
public class XmlToExcelWriterTest {
    private static final int ROWS = 10;
    private ExcelMetaData metadata;
    private List<Map<String, Object>> datas;

    /**
     * @throws java.lang.Exception
     */
    @Before
    public void setUp() throws Exception {
        metadata = new ExcelMetaData();
        metadata.setFileName("expense");
        metadata.setFileType("xlsx");
        metadata.setSheetName("expense");
        metadata.setHasSubTitle(true);

        this.metadata.setHasHeader(true);
        ExcelHeader header = new ExcelHeader();
        header.setHeaderName("报销单");
        this.metadata.setHeader(header);

        this.metadata.setHasFooter(true);
        ExcelFooter footer = new ExcelFooter();
        footer.setRemarks("说明:本文档版权所有,违法必究");
        this.metadata.setFooter(footer);

        structureMetaData();
        constructdata();
    }

    @Test
    public void test() {
        long begTime = System.currentTimeMillis();
        try {
            XmlToExcelWriter writer = new XmlToExcelWriter();
            writer.process(metadata, datas, "/Users/Robert/Desktop/QA_test/expense.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }

        System.out.println("Total:" + (System.currentTimeMillis() - begTime) + "ms");
    }

    @SuppressWarnings("unchecked")
    private void constructdata() {
        datas = new ArrayList<Map<String, Object>>();
        for (int i = 1; i <= ROWS; i++) {
            Expense e = new Expense("报销单" + i, "采购费用" + i, "Tim", "产品部", new Date(), "审批中", 880000 + i, 3200 + i);
            List<CostDetail> detail = new ArrayList<CostDetail>();
            CostDetail c1 = new CostDetail("酒店", new Date(), "培训酒店住宿", 3000 + i);
            detail.add(c1);
            CostDetail c2 = new CostDetail("酒店", new Date(), "培训酒店住宿", 32000 + i);
            detail.add(c2);
            CostDetail c3 = new CostDetail("酒店", new Date(), "培训酒店住宿", 35000 + i);
            detail.add(c3);
            CostDetail c4 = new CostDetail("酒店", new Date(), "培训酒店住宿", 8000 + i);
            detail.add(c4);
            CostDetail c5 = new CostDetail("餐费", new Date(), "培训午餐费", 800 + i);
            detail.add(c5);
            e.setCostDetail(detail);
            String es = JsonUtil.beanToString(e);
            Map<String, Object> em = JsonUtil.stringToBean(es, Map.class);
            datas.add(em);
        }
    }

    /**
     * 
     */
    private void structureMetaData() {
        // 设置表头
        List<ExcelTitle> titles = new ArrayList<ExcelTitle>();
        ExcelTitle t1 = new ExcelTitle();
        t1.setIndex(0);
        t1.setName("billType");
        t1.setDisplayName("单据类型");
        t1.setDataType(CellType.TEXT);
        t1.setMerge(true);
        titles.add(t1);

        ExcelTitle t2 = new ExcelTitle();
        t2.setIndex(1);
        t2.setName("billName");
        t2.setDisplayName("单据名称");
        t2.setDataType(CellType.TEXT);
        t2.setMerge(true);
        titles.add(t2);

        ExcelTitle t3 = new ExcelTitle();
        t3.setIndex(2);
        t3.setName("createUserId");
        t3.setDisplayName("提交人");
        t3.setDataType(CellType.TEXT);
        t3.setMerge(true);
        titles.add(t3);

        ExcelTitle t4 = new ExcelTitle();
        t4.setIndex(3);
        t4.setName("owner");
        t4.setDisplayName("费用归属");
        t4.setDataType(CellType.TEXT);
        t4.setMerge(true);
        titles.add(t4);

        ExcelTitle t5 = new ExcelTitle();
        t5.setIndex(4);
        t5.setName("submitDate");
        t5.setDisplayName("审批提交日期");
        t5.setDataType(CellType.DATE);
        t5.setMerge(true);
        titles.add(t5);

        ExcelTitle t6 = new ExcelTitle();
        t6.setIndex(5);
        t6.setName("status");
        t6.setDisplayName("审批状态");
        t6.setDataType(CellType.TEXT);
        t6.setMerge(true);
        titles.add(t6);

        ExcelTitle t7 = new ExcelTitle();
        t7.setName("costDetail");
        t7.setDisplayName("费用详情");
        t7.setDataType(CellType.LIST);
        t7.setHasSubTitle(true);
        titles.add(t7);

        List<ExcelTitle> subTitles = new ArrayList<ExcelTitle>();
        // 初始化子项
        ExcelTitle t8 = new ExcelTitle();
        t8.setIndex(6);
        t8.setName("costtype");
        t8.setDisplayName("费用类型");
        t8.setDataType(CellType.TEXT);
        subTitles.add(t8);

        ExcelTitle t9 = new ExcelTitle();
        t9.setIndex(7);
        t9.setName("costCreateTime");
        t9.setDisplayName("费用发生时间");
        t9.setDataType(CellType.DATE);
        subTitles.add(t9);

        ExcelTitle t10 = new ExcelTitle();
        t10.setIndex(8);
        t10.setName("costDesc");
        t10.setDisplayName("费用描述");
        t10.setDataType(CellType.TEXT);
        subTitles.add(t10);

        ExcelTitle t11 = new ExcelTitle();
        t11.setIndex(9);
        t11.setName("costMoney");
        t11.setDisplayName("费用金额");
        t11.setDataType(CellType.MONEY);
        subTitles.add(t11);

        t7.setSubTitles(subTitles);

        ExcelTitle t12 = new ExcelTitle();
        t12.setIndex(10);
        t12.setName("expenseMoney");
        t12.setDisplayName("报销金额");
        t12.setDataType(CellType.MONEY);
        t12.setMerge(true);
        titles.add(t12);

        ExcelTitle t13 = new ExcelTitle();
        t13.setIndex(11);
        t13.setName("loanMoney");
        t13.setDisplayName("借款金额");
        t13.setDataType(CellType.MONEY);
        t13.setMerge(true);
        titles.add(t13);

        metadata.setExcelTitle(titles);
    }

    class Expense {
        private String billType;
        private String billName;
        private String createUserId;
        private String owner;
        private Date submitDate;
        private String status;
        private List<CostDetail> costDetail;
        private double expenseMoney;
        private double loanMoney;

        public Expense(String billType, String billName, String createUserId, String owner, Date submitDate, String status, double expenseMoney, double loanMoney) {
            this.billType = billType;
            this.billName = billName;
            this.createUserId = createUserId;
            this.owner = owner;
            this.submitDate = submitDate;
            this.status = status;
            this.expenseMoney = expenseMoney;
            this.loanMoney = loanMoney;
        }

        public List<CostDetail> getCostDetail() {
            return costDetail;
        }

        public void setCostDetail(List<CostDetail> costDetail) {
            this.costDetail = costDetail;
        }

        public String getBillType() {
            return billType;
        }

        public void setBillType(String billType) {
            this.billType = billType;
        }

        public String getBillName() {
            return billName;
        }

        public void setBillName(String billName) {
            this.billName = billName;
        }

        public String getCreateUserId() {
            return createUserId;
        }

        public void setCreateUserId(String createUserId) {
            this.createUserId = createUserId;
        }

        public String getOwner() {
            return owner;
        }

        public void setOwner(String owner) {
            this.owner = owner;
        }

        public Date getSubmitDate() {
            return submitDate;
        }

        public void setSubmitDate(Date submitDate) {
            this.submitDate = submitDate;
        }

        public String getStatus() {
            return status;
        }

        public void setStatus(String status) {
            this.status = status;
        }

        public double getExpenseMoney() {
            return expenseMoney;
        }

        public void setExpenseMoney(double expenseMoney) {
            this.expenseMoney = expenseMoney;
        }

        public double getLoanMoney() {
            return loanMoney;
        }

        public void setLoanMoney(double loanMoney) {
            this.loanMoney = loanMoney;
        }
    }

    class CostDetail {
        private String costtype;
        private Date costCreateTime;
        private String costDesc;
        private double costMoney;

        public CostDetail(String costtype, Date costCreateTime, String costDesc, double costMoney) {
            this.costtype = costtype;
            this.costCreateTime = costCreateTime;
            this.costDesc = costDesc;
            this.costMoney = costMoney;
        }

        public String getCosttype() {
            return costtype;
        }

        public void setCosttype(String costtype) {
            this.costtype = costtype;
        }

        public Date getCostCreateTime() {
            return costCreateTime;
        }

        public void setCostCreateTime(Date costCreateTime) {
            this.costCreateTime = costCreateTime;
        }

        public String getCostDesc() {
            return costDesc;
        }

        public void setCostDesc(String costDesc) {
            this.costDesc = costDesc;
        }

        public double getCostMoney() {
            return costMoney;
        }

        public void setCostMoney(double costMoney) {
            this.costMoney = costMoney;
        }
    }
}