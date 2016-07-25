package com.xuyihao.java2excel;

import com.xuyihao.java2excel.api.model.AttributeType;
import com.xuyihao.java2excel.api.model.ExcelTemplate;
import com.xuyihao.java2excel.biz.excel.exportfunc.ExportUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/25 下午 03:55.
 */
public class JohnsonTestMain {
    public static void main(String args[]){
        try{
            System.out.println(createExcelTest());
            System.out.println(insertExcelDataTest());
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static boolean createExcelTest() throws Exception {
        ExportUtil exportUtil = new ExportUtil();
        FileOutputStream fs;
        File file = new File("E:\\JUnitTestPath\\testTemplate.xls");
        fs = new FileOutputStream(file);
        Workbook workbook = new HSSFWorkbook();
        int sheetNum = 0;
        ExcelTemplate template = new ExcelTemplate();
        template.setTenant("adhauihdiuahfd");
        template.setId("13415641");
        template.setClassCode("Y_Application");
        template.setClassName("应用");
        List<AttributeType> attributeTypes = new ArrayList<AttributeType>();
        setAttributes(attributeTypes);
        template.setAttrbuteTypes(attributeTypes);
        boolean flag = exportUtil.createExcel(workbook, sheetNum, template, true, fs);
        fs.close();
        return flag;
    }

    public static boolean insertExcelDataTest() throws Exception{
        ExportUtil exportUtil = new ExportUtil();
        File file = new File("E:\\JUnitTestPath\\testTemplate.xls");
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new HSSFWorkbook(fis);
        int sheetNum = 0;
        List<ExcelTemplate> datas = new ArrayList<ExcelTemplate>();
        setValues(datas);
        boolean flag = exportUtil.insertExcelData(workbook, sheetNum, 6, datas, true, new FileOutputStream(file));
        return flag;
    }

    public static void setAttributes(List<AttributeType> attributeTypes){
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
        attributeTypes.add(attribute2);

        AttributeType attribute3 = new AttributeType();
        attribute3.setAttrCode("Y_School");
        attribute3.setAttrFormatRule("yeah");
        attribute3.setAttrId("416584");
        attribute3.setAttrName("学校");
        attribute3.setAttrType("Text");
        attribute3.setDefaultValue("122");
        attributeTypes.add(attribute3);

        AttributeType attribute4 = new AttributeType();
        attribute4.setAttrCode("Y_age");
        attribute4.setAttrFormatRule("yeah");
        attribute4.setAttrId("45468");
        attribute4.setAttrName("年龄");
        attribute4.setAttrType("Text");
        attribute4.setDefaultValue("122");
        attributeTypes.add(attribute4);
    }

    public static void setValues(List<ExcelTemplate> excelTemplates){
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
        template2.setTenant("adhauihdiuahfd");
        template2.setId("13415641");
        template2.setClassCode("Y_Application");
        template2.setClassName("应用");
        List<String> data2 = template2.getAttrValues();
        data2.add("李四");
        data2.add("杭电");
        data2.add("201325414874");
        data2.add("20");
        excelTemplates.add(template2);

        ExcelTemplate template3 = new ExcelTemplate();
        template3.setTenant("adhauihdiuahfd");
        template3.setId("13415641");
        template3.setClassCode("Y_Application");
        template3.setClassName("应用");
        List<String> data3 = template3.getAttrValues();
        data3.add("王五");
        data3.add("浙大");
        data3.add("201654241111");
        data3.add("18");
        excelTemplates.add(template3);

        ExcelTemplate template4 = new ExcelTemplate();
        template4.setTenant("adhauihdiuahfd");
        template4.setId("13415641");
        template4.setClassCode("Y_Application");
        template4.setClassName("应用");
        List<String> data4 = template4.getAttrValues();
        data4.add("李二");
        data4.add("工大");
        data4.add("214452360124");
        data4.add("22");
        excelTemplates.add(template4);
    }
}
