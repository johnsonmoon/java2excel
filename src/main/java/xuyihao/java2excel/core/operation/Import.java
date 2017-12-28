package xuyihao.java2excel.core.operation;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Template;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/27 下午 12:48.
 * <p>
 * Excel文件操作工具类（读取excel数据）
 */
public class Import {
    private static Logger logger = LoggerFactory.getLogger(Import.class);

    /**
     * 从传入的workbook指定的sheet中获取ExcelTemplate对象
     *
     * @param workbook    表格
     * @param sheetNumber 工作簿编号
     * @return ExcelTemplate对象
     */
    public static Template getExcelTemplateFromExcel(Workbook workbook, int sheetNumber) {
        Template template = new Template();
        Sheet sheet = workbook.getSheetAt(sheetNumber);
        String headValue = Common.getCellValue(sheet, 0, 0);
        template.setName(headValue);
        List<Attribute> attributeList = new ArrayList<>();
        int attrCount = sheet.getRow(3).getPhysicalNumberOfCells() - 2;
        for (int i = 0; i < attrCount; i++) {
            String[] attributeTypeInfo = Common.getCellValue(sheet, 3, i + 2).split("&&");
            Attribute attribute = new Attribute();
            attribute.setAttrCode(attributeTypeInfo[0]);
            attribute.setAttrName(attributeTypeInfo[1]);
            attribute.setAttrType(attributeTypeInfo[2]);
            attributeList.add(attribute);
        }
        template.setAttributes(attributeList);
        return template;
    }

    /**
     * 从传入的workbook指定的sheet中获取具体数据条数（记录条数）
     *
     * @param workbook    表格
     * @param sheetNumber 工作簿编号
     * @return 记录条数
     */
    public static int getAttrValueCount(Workbook workbook, int sheetNumber) {
        int totalValueCount = 0;
        int count = workbook.getSheetAt(sheetNumber).getLastRowNum();
        if (count == 6) {
            String checkCellValue = Common.getCellValue(workbook.getSheetAt(sheetNumber), 6, 2);
            if (checkCellValue == null || checkCellValue.equals("")) {
                totalValueCount = 0;
            }
        } else {
            totalValueCount = count - 6 + 1;
        }
        return totalValueCount;
    }

    /**
     * 从excel表格中获取具体的数据，返回一个列表
     *
     * @param workbook    表格
     * @param sheetNumber 工作簿编号
     * @param beginRow    读取起始行
     * @param readSize    一次读取的行数
     * @return 具体数据列表
     */
    public static List<Template> getExcelTemplateListDataFromExcel(Workbook workbook, int sheetNumber, int beginRow,
                                                                   int readSize) {
        List<Template> templateList = new ArrayList<>();
        if (beginRow < 6) {
            logger.warn("Insert position must above 7(6+1)!");
            return templateList;
        }
        Template template = getExcelTemplateFromExcel(workbook, sheetNumber);
        Sheet sheet = workbook.getSheetAt(sheetNumber);
        for (int i = 0; i < readSize; i++) {
            String value = Common.getCellValue(sheet, i + beginRow, 2);
            if ((value == null) || (value.equals(""))) {
                break;
            }
            Template data = new Template();
            data.setName(template.getName());
            data.setAttributes(template.getAttributes());
            List<String> attrValues = new ArrayList<>();
            for (int j = 0; j < data.getAttributes().size(); j++) {
                attrValues.add(Common.getCellValue(sheet, i + beginRow, j + 2));
            }
            data.setAttrValues(attrValues);
            templateList.add(data);
        }
        return templateList;
    }
}
