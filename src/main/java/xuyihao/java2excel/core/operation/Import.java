package xuyihao.java2excel.core.operation;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Model;
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
	public static Model getTemplateFromExcel(Workbook workbook, int sheetNumber) {
		if (workbook == null)
			return null;
		if (sheetNumber < 0)
			sheetNumber = 0;
		Model model = new Model();
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		String headValue = Common.getCellValue(sheet, 0, 0);
		model.setJavaClassName(headValue);
		String nameValue = Common.getCellValue(sheet, 1, 0);
		model.setName(nameValue);
		int attrCount = sheet.getRow(3).getPhysicalNumberOfCells() - 2;
		for (int i = 0; i < attrCount; i++) {
			String[] attributeTypeInfo = Common.getCellValue(sheet, i + 2, 3).split("&&");
			Attribute attribute = new Attribute();
			attribute.setAttrCode(replaceNull(attributeTypeInfo[0]));
			attribute.setAttrName(replaceNull(attributeTypeInfo[1]));
			attribute.setAttrType(replaceNull(attributeTypeInfo[2]));
			attribute.setFormatInfo(replaceNull(attributeTypeInfo[3]));
			attribute.setDefaultValue(replaceNull(attributeTypeInfo[4]));
			attribute.setUnit(replaceNull(attributeTypeInfo[5]));
			attribute.setJavaClassName(replaceNull(attributeTypeInfo[6]));
			model.addAttribute(attribute);
		}
		return model;
	}

	private static String replaceNull(String arrayValue) {
		if (arrayValue == null)
			return null;
		if (arrayValue.isEmpty())
			return null;
		if (arrayValue.equals("null"))
			return null;
		return arrayValue;
	}

	/**
	 * 从传入的workbook指定的sheet中获取具体数据条数（记录条数）
	 *
	 * @param workbook    表格
	 * @param sheetNumber 工作簿编号
	 * @return 记录条数
	 */
	public static int getDataCountFromExcel(Workbook workbook, int sheetNumber) {
		if (workbook == null)
			return 0;
		if (sheetNumber < 0)
			sheetNumber = 0;
		int totalValueCount = 0;
		int count = workbook.getSheetAt(sheetNumber).getLastRowNum();
		if (count == 6) {
			String checkCellValue = Common.getCellValue(workbook.getSheetAt(sheetNumber), 2, 6);
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
	 * @param beginRow    读取起始行（begin from 0）
	 * @param readSize    一次读取的行数
	 * @return 具体数据列表
	 */
	public static List<Model> getDataFromExcel(Workbook workbook, int sheetNumber, int beginRow,
			int readSize) {
		List<Model> modelList = new ArrayList<>();
		if (workbook == null)
			return modelList;
		if (sheetNumber < 0)
			sheetNumber = 0;
		if (beginRow < 0) {
			logger.warn("Read position must above 0!");
			return modelList;
		}
		if (readSize < 0) {
			logger.warn("Read size must not below 0!");
			return modelList;
		}
		Model model = getTemplateFromExcel(workbook, sheetNumber);
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		for (int i = 0; i < readSize; i++) {
			String value = Common.getCellValue(sheet, 2, i + beginRow + 6);
			if ((value == null) || (value.equals(""))) {
				break;
			}
			Model data = new Model();
			data.setName(model.getName());
			data.setJavaClassName(model.getJavaClassName());
			data.setAttributes(model.getAttributes());
			List<String> attrValues = new ArrayList<>();
			for (int j = 0; j < data.getAttributes().size(); j++) {
				attrValues.add(Common.getCellValue(sheet, j + 2, i + beginRow + 6));
			}
			data.setAttrValues(attrValues);
			modelList.add(data);
		}
		return modelList;
	}
}
