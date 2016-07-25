package com.xuyihao.java2excel.test.biz.excel.exportfunc;

import com.xuyihao.java2excel.api.model.AttributeType;
import com.xuyihao.java2excel.api.model.ExcelTemplate;
import com.xuyihao.java2excel.biz.excel.exportfunc.ExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/25 下午 01:39.
 */
public class exportUtilTest {

    ExportUtil exportUtil = new ExportUtil();

    private void setAttributes(List<AttributeType> attributeTypes){
        AttributeType attribute1 = new AttributeType();
        attribute1.setAttrCode("Y_name");
        attribute1.setAttrFormatRule("yeah");
        attribute1.setAttrId("123");
        attribute1.setAttrName("姓名");
        attribute1.setAttrType("Text");
        attribute1.setDefaultValue("122");
        attributeTypes.add(attribute1);

        AttributeType attribute2 = new AttributeType();
        attribute2.setAttrCode("Y_num");
        attribute2.setAttrFormatRule("yeah");
        attribute2.setAttrId("54785");
        attribute2.setAttrName("学号");
        attribute2.setAttrType("Text");
        attribute2.setDefaultValue("122");
        attributeTypes.add(attribute1);

        AttributeType attribute3 = new AttributeType();
        attribute3.setAttrCode("Y_School");
        attribute3.setAttrFormatRule("yeah");
        attribute3.setAttrId("416584");
        attribute3.setAttrName("学校");
        attribute3.setAttrType("Text");
        attribute3.setDefaultValue("122");
        attributeTypes.add(attribute1);

        AttributeType attribute4 = new AttributeType();
        attribute4.setAttrCode("Y_age");
        attribute4.setAttrFormatRule("yeah");
        attribute4.setAttrId("45468");
        attribute4.setAttrName("年龄");
        attribute4.setAttrType("Text");
        attribute4.setDefaultValue("122");
        attributeTypes.add(attribute1);
    }

    private void setValues(List<ExcelTemplate> excelTemplates){
        ExcelTemplate template = new ExcelTemplate();
        template.setTenant("adhauihdiuahfd");
        template.setId("13415641");
        template.setClassCode("Y_Application");
        template.setClassName("应用");
        List<String> data = template.getAttrValues();
        data.add("张三");
        data.add("201326810966");
        data.add("工大");
        data.add("19");
        excelTemplates.add(template);

        ExcelTemplate template2 = new ExcelTemplate();
        template.setTenant("adhauihdiuahfd");
        template.setId("13415641");
        template.setClassCode("Y_Application");
        template.setClassName("应用");
        List<String> data2 = template.getAttrValues();
        data.add("李四");
        data.add("杭电");
        data.add("201325414874");
        data.add("20");
        excelTemplates.add(template2);

        ExcelTemplate template3 = new ExcelTemplate();
        template.setTenant("adhauihdiuahfd");
        template.setId("13415641");
        template.setClassCode("Y_Application");
        template.setClassName("应用");
        List<String> data3 = template.getAttrValues();
        data.add("王五");
        data.add("浙大");
        data.add("201654241111");
        data.add("18");
        excelTemplates.add(template3);

        ExcelTemplate template4 = new ExcelTemplate();
        template.setTenant("adhauihdiuahfd");
        template.setId("13415641");
        template.setClassCode("Y_Application");
        template.setClassName("应用");
        List<String> data4 = template.getAttrValues();
        data.add("李二");
        data.add("工大");
        data.add("214452360124");
        data.add("22");
        excelTemplates.add(template4);
    }

    @Test(expected = IOException.class)
    public void createExcelTest() throws Exception {
        FileOutputStream fs;
        File file = new File("E:\\JUnitTestPath\testTemplate.xls");
        fs = new FileOutputStream(file);
        Workbook workbook = new HSSFWorkbook();
        int sheetNum = 0;
        ExcelTemplate template = new ExcelTemplate();
        template.setTenant("adhauihdiuahfd");
        template.setId("13415641");
        template.setClassCode("Y_Application");
        template.setClassName("应用");
        List<AttributeType> attributeTypes = new ArrayList<AttributeType>();
        this.setAttributes(attributeTypes);
        template.setAttrbuteTypes(attributeTypes);
        boolean flag = exportUtil.createExcel(workbook, sheetNum, template, false, null);
        workbook.write(fs);
        fs.close();
        workbook.close();
        Assert.assertEquals(true, flag);
    }

    @Test
    public void insertExcelDataTest() throws Exception{
        File file = new File("E:\\JUnitTestPath\testTemplate.xls");
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new HSSFWorkbook(fis);
        int sheetNum = 0;
        List<ExcelTemplate> datas = new ArrayList<ExcelTemplate>();
        this.setValues(datas);
        boolean flag = exportUtil.insertExcelData(workbook, sheetNum, 6, datas, true, new FileOutputStream(file));
        Assert.assertEquals(true, flag);
    }
}
