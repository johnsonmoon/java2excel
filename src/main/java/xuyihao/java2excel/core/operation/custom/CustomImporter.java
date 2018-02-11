package xuyihao.java2excel.core.operation.custom;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.custom.map.ColumnMapper;
import xuyihao.java2excel.core.operation.Common;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/2/11 12:27.
 */
public class CustomImporter {
	private static Logger logger = LoggerFactory.getLogger(CustomImporter.class);

	/**
	 * Get data count from excel workbook sheet.
	 *
	 * @param workbook       excel workbook
	 * @param sheetNumber    sheet number
	 * @param beginRowNumber row number where data begin at (start from 0)
	 * @return data count
	 */
	public static int getDataCountFromExcel(Workbook workbook, int sheetNumber, int beginRowNumber) {
		if (workbook == null)
			return 0;
		if (sheetNumber < 0)
			sheetNumber = 0;
		if (beginRowNumber < 0)
			return 0;
		int totalValueCount;
		int lastRowNum = workbook.getSheetAt(sheetNumber).getLastRowNum();
		if (beginRowNumber > lastRowNum) {
			totalValueCount = 0;
		} else {
			totalValueCount = lastRowNum - beginRowNumber + 1;
		}
		return totalValueCount;
	}

	/**
	 * Get data list from sheet.
	 * <p>
	 * <pre>
	 *     data map:
	 *     [
	 *          {columnNumber:columnValue},
	 *          {columnNumber:columnValue},
	 *          {columnNumber:columnValue}
	 *     ]
	 * </pre>
	 *
	 * @param workbook    excel workbook
	 * @param sheetNumber sheet number
	 * @param beginRow    row read begin at（start from 0）
	 * @param readSize    row count to read
	 * @return map data list
	 */
	public static List<Map<Integer, Object>> getDataFromExcel(Workbook workbook, int sheetNumber, int beginRow,
			int readSize) {
		List<Map<Integer, Object>> dataMapList = new ArrayList<>();
		if (workbook == null)
			return dataMapList;
		if (sheetNumber < 0)
			sheetNumber = 0;
		if (beginRow < 0) {
			logger.warn("Read position must above 0!");
			return dataMapList;
		}
		int lastRowNum = workbook.getSheetAt(sheetNumber).getLastRowNum();
		if (beginRow > lastRowNum) {
			return dataMapList;
		}
		if (readSize < 0) {
			logger.warn("Read size must not below 0!");
			return dataMapList;
		}
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		for (int row = 0; row < readSize; row++) {
			int currentRow = row + beginRow;
			if (currentRow > lastRowNum)
				break;
			Map<Integer, Object> rowData = new HashMap<>();
			int column = 0;
			String value;
			do {
				value = Common.getCellValue(sheet, column, currentRow);
				if (value != null && !value.isEmpty()) {
					rowData.put(column, value);
				}
				column++;
			} while (value != null && !value.isEmpty());
			dataMapList.add(rowData);
		}
		return dataMapList;
	}

	/**
	 * Get data list from sheet.
	 * <p>
	 * <pre>
	 *     data map:
	 *     [
	 *          {fieldName:columnValue},
	 *          {fieldName:columnValue},
	 *          {fieldName:columnValue}
	 *     ]
	 * </pre>
	 *
	 * @param workbook     excel workbook
	 * @param sheetNumber  sheet number
	 * @param columnMapper mapping for columnNumber-fieldName
	 * @param beginRow     row read begin at（start from 0）
	 * @param readSize     row count to read
	 * @return map data list
	 */
	public static List<Map<String, Object>> getDataFromExcel(Workbook workbook, int sheetNumber,
			ColumnMapper columnMapper, int beginRow, int readSize) {
		List<Map<String, Object>> dataMapList = new ArrayList<>();
		if (workbook == null)
			return dataMapList;
		if (sheetNumber < 0)
			sheetNumber = 0;
		if (columnMapper == null || columnMapper.getOrderedMaps(ColumnMapper.ORDER_TYPE_ASC) == null) {
			logger.warn("Column mapping data must not be null!");
			return dataMapList;
		}
		if (beginRow < 0) {
			logger.warn("Read position must above 0!");
			return dataMapList;
		}
		int lastRowNum = workbook.getSheetAt(sheetNumber).getLastRowNum();
		if (beginRow > lastRowNum) {
			return dataMapList;
		}
		if (readSize < 0) {
			logger.warn("Read size must not below 0!");
			return dataMapList;
		}
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		for (int row = 0; row < readSize; row++) {
			int currentRow = row + beginRow;
			if (currentRow > lastRowNum)
				break;
			Map<String, Object> rowData = new HashMap<>();
			for (ColumnMapper.Map map : columnMapper.getOrderedMaps(ColumnMapper.ORDER_TYPE_ASC)) {
				Integer column = map.columnNumber;
				String fieldName = map.fieldName;
				if (column == null || fieldName == null || fieldName.isEmpty())
					continue;
				String value = Common.getCellValue(sheet, column, currentRow);
				rowData.put(fieldName, value);
			}
			dataMapList.add(rowData);
		}
		return dataMapList;
	}
}
