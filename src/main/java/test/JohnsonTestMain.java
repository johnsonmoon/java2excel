package test;

import org.apache.poi.ss.usermodel.Sheet;
import xuyihao.java2excel.excel.util.CommonExcelUtil;
import xuyihao.java2excel.excel.util.ImportUtil;
import xuyihao.java2excel.excel.entity.AttributeType;
import xuyihao.java2excel.excel.entity.ExcelTemplate;
import xuyihao.java2excel.excel.util.ExportUtil;
import xuyihao.java2excel.excel.entity.ProgressMessage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/25 下午 03:55.
 */
public class JohnsonTestMain {
    public static void main(String args[]){
        try{
            //testGenerateWarningMessage();
            //getExcelTemplateListDataFromExcel();
            //testGetAttrValueCount();
            //testReadExcelTemplateFromExcel();
            //System.out.println(createExcelTest(new ProgressMessage()));
            //System.out.println(insertExcelDataTest(new ProgressMessage()));
            //System.out.println(CurrentTimeUtil.getCurrentTime());
            //createExcelTest(new ProgressMessage());
            testGetCellValue();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void testGetCellValue()throws Exception{
        File file = new File("C:\\Users\\Administrator\\Desktop\\number_Template.xlsx");
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        ProgressMessage progressMessage = new ProgressMessage();
        ImportUtil importUtil = new ImportUtil();
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 6; i <= 8; i++) {
            String value = CommonExcelUtil.getCellValue(sheet, i, 5);
            System.out.println(value);
        }
    }

    public static void testGenerateWarningMessage() throws FileNotFoundException{
        File file = new File("E:\\JUnitTestPath\\WarningMessageExcelTest.xlsx");
        FileOutputStream outputStream = new FileOutputStream(file);
        ProgressMessage progressMessage = new ProgressMessage();
        List<String> warningMessageList = new ArrayList<String>();
        for(int i = 0; i < 100; i++){
            warningMessageList.add("ifahouh" + String.valueOf(i+8));
        }
        progressMessage.setWarningMessageList(warningMessageList);
    }

    public static void getExcelTemplateListDataFromExcel() throws Exception{
        File file = new File("E:\\JUnitTestPath\\MultiSheetTest.xlsx");
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        ProgressMessage progressMessage = new ProgressMessage();
        ImportUtil importUtil = new ImportUtil();
        int count = importUtil.getAttrValueCount(workbook, 12, progressMessage);
        for(int i = 6; i < count; i += 12){
            List<ExcelTemplate> excelTemplateList = importUtil.getExcelTemplateListDataFromExcel(workbook, 12, i, 12, progressMessage);
            for(int j = 0; j < excelTemplateList.size(); j++){
                String valueAll = "";
                for(int k = 0; k < excelTemplateList.get(j).getAttrValues().size(); k++){
                    valueAll = valueAll + (" && " + excelTemplateList.get(j).getAttrValues().get(k).trim());
                }
                System.out.println(excelTemplateList.get(j).getClassName() + ": " + valueAll);
            }
        }
    }

    public static void testGetAttrValueCount()throws Exception{
        File file = new File("E:\\JUnitTestPath\\MultiSheetTest.xlsx");
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        ProgressMessage progressMessage = new ProgressMessage();
        ImportUtil importUtil = new ImportUtil();
        for(int i = 0 ; i < workbook.getNumberOfSheets(); i ++){
            ExcelTemplate template = importUtil.getExcelTemplateFromExcel(workbook, i, progressMessage);
            int dataCount = importUtil.getAttrValueCount(workbook, i, progressMessage);
            System.out.println(template.getClassCode() + "  " + template.getClassName() + ":" + dataCount);
        }
    }

    public static void testReadExcelTemplateFromExcel() throws Exception{
        File file = new File("E:\\JUnitTestPath\\test.xlsx");
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        ProgressMessage progressMessage = new ProgressMessage();
        ImportUtil importUtil = new ImportUtil();
        ExcelTemplate template = importUtil.getExcelTemplateFromExcel(workbook, 0, progressMessage);
        System.out.println(template.getTenant());
        System.out.println(template.getClassCode());
        System.out.println(template.getClassName());
        List<AttributeType> attributeTypeList = template.getAttrbuteTypes();
        for(int i = 0; i < attributeTypeList.size(); i++){
            System.out.println(attributeTypeList.get(i).getAttrId() + "&&" + attributeTypeList.get(i).getAttrCode() + "&&" + attributeTypeList.get(i).getAttrName());
        }
    }

    public static void testProgressMessage(){
        ProgressMessage progressMessage = new ProgressMessage();
        //重置
        progressMessage.reset();
        progressMessage.setTotalCount(500);
        progressMessage.setDetailMessage("共有 500条 操作，开始计算");
        progressMessage.stateNotStart();
        //开始计算
        progressMessage.stateStarted();
        for(int i = 0; i <= 50000; i++) {
            progressMessage.addCurrentCount();
        }
        progressMessage.stateEnd();
        progressMessage.setDetailMessage("操作结束！");
    }

    public static boolean createExcelTest(ProgressMessage progressMessage) throws Exception {
        ExportUtil exportUtil = new ExportUtil();
        FileOutputStream fs;
        File file = new File("E:\\JUnitTestPath\\test.xlsx");
        fs = new FileOutputStream(file);
        Workbook workbook = new XSSFWorkbook();
        int sheetNum = 0;
        ExcelTemplate template = new ExcelTemplate();
        template.setTenant("adhauihdiuahfd");
        template.setClassCode("Y_Application");
        template.setClassName("应用");
        List<AttributeType> attributeTypes = new ArrayList<AttributeType>();
        setAttributes(attributeTypes);
        template.setAttrbuteTypes(attributeTypes);
        boolean flag = exportUtil.createExcel(workbook, sheetNum, template, true, fs, progressMessage);
        fs.close();
        return flag;
    }

    public static boolean insertExcelDataTest(ProgressMessage progressMessage) throws Exception{
        ExportUtil exportUtil = new ExportUtil();
        File file = new File("E:\\JUnitTestPath\\test.xlsx");
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        int sheetNum = 0;
        List<ExcelTemplate> datas = new ArrayList<ExcelTemplate>();
        setValues(datas);
        boolean flag2 = exportUtil.insertExcelData(workbook, sheetNum, 6, datas.get(0), datas, true, new FileOutputStream(file), progressMessage);
        return flag2;
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
