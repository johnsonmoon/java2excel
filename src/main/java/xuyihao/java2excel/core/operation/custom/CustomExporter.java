package xuyihao.java2excel.core.operation.custom;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.custom.map.ColumnMapper;
import xuyihao.java2excel.core.operation.Common;

import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/2/11 12:27.
 */
public class CustomExporter {
	private static Logger logger = LoggerFactory.getLogger(CustomExporter.class);

	/**
	 * Create excel sheet with column mapping, write table head with columnMapper data at row 0 (beginning row of the sheet).
	 *
	 * @param workbook     excel workbook
	 * @param sheetNum     sheet number
	 * @param sheetName    name for the sheet
	 * @param columnMapper mapping for columnNumber-fieldName
	 * @return true/false
	 */
	public static boolean createExcel(final Workbook workbook, int sheetNum, String sheetName,
			ColumnMapper columnMapper) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (columnMapper == null || columnMapper.getOrderedMaps(ColumnMapper.ORDER_TYPE_ASC) == null)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, sheetName);
			// column width
			int columnSize = columnMapper.getColumnNumbers().size();
			for (int i = 0; i < columnSize; i++) {
				sheet.setColumnWidth(i, 4800);
			}
			for (ColumnMapper.Map map : columnMapper.getOrderedMaps(ColumnMapper.ORDER_TYPE_ASC)) {
				int columnNumber = map.columnNumber;
				String fieldName = map.fieldName;
				if (fieldName == null || fieldName.isEmpty())
					continue;
				Common.insertCellValue(sheet, columnNumber, 0, fieldName,
						Common.createCellStyle(workbook, Common.CELL_STYLE_TYPE_ROW_HEADER));
			}
			flag = true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * insert data
	 *
	 * @param workbook    excel workbook
	 * @param sheetNum    sheet number
	 * @param startRowNum row number to start writing (begin from 0)
	 * @param dataMapList map data list to write with
	 *                    [
	 *                    {columnNumber:value},
	 *                    {columnNumber:value}
	 *                    ]
	 * @return true/false
	 */
	public static boolean insertExcelData(final Workbook workbook, int sheetNum, int startRowNum,
			List<Map<Integer, Object>> dataMapList) {
		boolean flag;
		try {
			if (dataMapList == null || dataMapList.isEmpty())
				return false;
			if (workbook == null) {
				return false;
			}
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			if (sheet == null)
				return false;
			if (startRowNum < 0) {
				logger.warn("Data insert must above 0!");
			}
			for (int i = 0; i < dataMapList.size(); i++) {
				Map<Integer, Object> map = dataMapList.get(i);
				int row = i + startRowNum;
				for (Map.Entry<Integer, Object> entry : map.entrySet()) {
					int column = entry.getKey();
					String value = entry.getValue() == null ? "" : String.valueOf(entry.getValue());
					Common.insertCellValue(sheet, column, row, value,
							Common.createCellStyle(workbook, Common.CELL_STYLE_TYPE_VALUE));
				}
			}
			flag = true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * insert data
	 *
	 * @param workbook     excel workbook
	 * @param sheetNum     sheet number
	 * @param startRowNum  row number to start writing (begin from 0)
	 * @param columnMapper mapping for columnNumber-fieldName
	 * @param dataMapList  map data list to write with
	 *                     [
	 *                     {fieldName:value},
	 *                     {fieldName:value}
	 *                     ]
	 * @return true/false
	 */
	public static boolean insertExcelData(final Workbook workbook, int sheetNum, int startRowNum,
			ColumnMapper columnMapper, List<Map<String, Object>> dataMapList) {
		boolean flag;
		try {
			if (dataMapList == null || dataMapList.isEmpty())
				return false;
			if (workbook == null) {
				return false;
			}
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			if (sheet == null)
				return false;
			if (startRowNum < 0) {
				logger.warn("Data insert must above 0!");
			}
			for (int i = 0; i < dataMapList.size(); i++) {
				Map<String, Object> dataMap = dataMapList.get(i);
				int row = i + startRowNum;
				for (ColumnMapper.Map columnMap : columnMapper.getOrderedMaps(ColumnMapper.ORDER_TYPE_ASC)) {
					int column = columnMap.columnNumber;
					String value = dataMap.get(columnMap.fieldName) == null ? ""
							: String.valueOf(dataMap.get(columnMap.fieldName));
					Common.insertCellValue(sheet, column, row, value,
							Common.createCellStyle(workbook, Common.CELL_STYLE_TYPE_VALUE));
				}
			}
			flag = true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}
}
