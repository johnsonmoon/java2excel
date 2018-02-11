package xuyihao.java2excel.core.operation.formatted;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Model;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.operation.Common;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/27 下午 12:48.
 */
public class FormattedImporter {
	private static Logger logger = LoggerFactory.getLogger(FormattedImporter.class);

	/**
	 * Getting model info from excel workbook sheet.
	 *
	 * @param workbook    excel workbook
	 * @param sheetNumber sheet number
	 * @return model instance
	 */
	public static Model getModelFromExcel(Workbook workbook, int sheetNumber) {
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
	 * Get data count from excel workbook sheet.
	 *
	 * @param workbook    excel workbook
	 * @param sheetNumber sheet number
	 * @return data count
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
	 * get model data from sheet
	 *
	 * @param workbook    excel workbook
	 * @param sheetNumber sheet number
	 * @param beginRow    row read begin at（begin from 0）
	 * @param readSize    row count to read
	 * @return model data
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
		Model model = getModelFromExcel(workbook, sheetNumber);
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
