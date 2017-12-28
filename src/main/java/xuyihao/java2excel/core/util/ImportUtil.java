package xuyihao.java2excel.core.util;

import xuyihao.java2excel.core.model.AttributeType;
import xuyihao.java2excel.core.model.ExcelTemplate;
import xuyihao.java2excel.core.model.ProgressMessage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/27 下午 12:48.
 *
 * Excel文件操作工具类（读取excel数据）
 */
public class ImportUtil{
    /**
     * 从传入的workbook指定的sheet中获取ExcelTemplate对象
     *
     * @param workbook 表格
     * @param sheetNumber 工作簿编号
     * @param progressMessage 进度信息
     * @return ExcelTemplate对象
     */
    public static ExcelTemplate getExcelTemplateFromExcel(Workbook workbook, int sheetNumber, ProgressMessage progressMessage) {
        ExcelTemplate excelTemplate = new ExcelTemplate();
        progressMessage.setDetailMessage("Analyzing sheet  " + sheetNumber + " .....");
        Sheet sheet = workbook.getSheetAt(sheetNumber);
        String[] headValue = CommonExcelUtil.getCellValue(sheet, 0, 0).split("&&");
        excelTemplate.setId(headValue[0]);
        excelTemplate.setTenant(headValue[1]);
        excelTemplate.setClassCode(headValue[2]);
        excelTemplate.setClassName(headValue[3]);
        List<AttributeType> attributeTypeList = excelTemplate.getAttrbuteTypes();
        int attrCount = sheet.getRow(3).getPhysicalNumberOfCells() - 2;
        for (int i = 0; i < attrCount; i++) {
            String[] attributeTypeInfo = CommonExcelUtil.getCellValue(sheet, 3, i + 2).split("&&");
            AttributeType attributeType = new AttributeType();
            attributeType.setAttrId(attributeTypeInfo[0]);
            attributeType.setAttrCode(attributeTypeInfo[1]);
            attributeType.setAttrName(attributeTypeInfo[2]);
            attributeTypeList.add(attributeType);
        }
        progressMessage.setDetailMessage("Class name: " + excelTemplate.getClassName());
        return excelTemplate;
    }

    /**
     * 从传入的workbook指定的sheet中获取具体数据条数（记录条数）
     *
     * @param workbook 表格
     * @param sheetNumber 工作簿编号
     * @param progressMessage 进度信息
     * @return 记录条数
     */
    public static int getAttrValueCount(Workbook workbook, int sheetNumber, ProgressMessage progressMessage) {
        int totalValueCount = 0;
        progressMessage.setDetailMessage("Analyzing sheet  " + sheetNumber + " .....");
        int count = workbook.getSheetAt(sheetNumber).getLastRowNum();
        if (count == 6) {
            String checkCellValue = CommonExcelUtil.getCellValue(workbook.getSheetAt(sheetNumber), 6, 2);
            if (checkCellValue == null || checkCellValue.equals("")) {
                totalValueCount = 0;
            }
        } else {
            totalValueCount = count - 6 + 1;
        }
        progressMessage.setDetailMessage("Total value count: " + totalValueCount);
        return totalValueCount;
    }

    /**
     * 从excel表格中获取具体的数据，返回一个列表
     *
     * @param workbook 表格
     * @param sheetNumber 工作簿编号
     * @param beginRow 读取起始行
     * @param readSize 一次读取的行数
     * @param progressMessage 进度信息
     * @return 具体数据列表
     */
    public static List<ExcelTemplate> getExcelTemplateListDataFromExcel(Workbook workbook, int sheetNumber, int beginRow,
                                                                 int readSize, ProgressMessage progressMessage) {
        List<ExcelTemplate> excelTemplateList = new ArrayList<ExcelTemplate>();
        ExcelTemplate template = getExcelTemplateFromExcel(workbook, sheetNumber, progressMessage);
        Sheet sheet = workbook.getSheetAt(sheetNumber);
        progressMessage.setDetailMessage("Read begin row: " + beginRow + " Read size" + readSize);
        if (beginRow < 6) {
            System.out.println("Insert position must above 7(6+1)!");
            progressMessage.stateFailed();
        }
        for (int i = 0; i < readSize; i++) {
            String value = CommonExcelUtil.getCellValue(sheet, i + beginRow, 2);
            if ((value == null) || (value.equals(""))) {
                break;
            }
            ExcelTemplate template1 = new ExcelTemplate(template);
            List<String> attrValues = new ArrayList<String>();
            for (int j = 0; j < template1.getAttrbuteTypes().size(); j++) {
                attrValues.add(CommonExcelUtil.getCellValue(sheet, i + beginRow, j + 2));
            }
            template1.setAttrValues(attrValues);
            excelTemplateList.add(template1);
        }
        return excelTemplateList;
    }
}
