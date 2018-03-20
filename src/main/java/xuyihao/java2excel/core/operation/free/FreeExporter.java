package xuyihao.java2excel.core.operation.free;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.free.CellData;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.util.ValueUtils;

import java.util.List;

/**
 * Created by xuyh at 2018/3/20 15:02.
 */
public class FreeExporter {
	private static Logger logger = LoggerFactory.getLogger(FreeExporter.class);

	/**
	 * Set column width for workbook at sheetNum column.
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param column   given column number
	 * @param width    width
	 * @return true/false
	 */
	public static boolean setColumnWidth(final Workbook workbook, int sheetNum, int column, int width) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (sheetNum < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			sheet.setColumnWidth(column, width);
			flag = true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Create excel sheet with name {sheetName}, with data {cellDataList}.
	 *
	 * @param workbook     excel workbook
	 * @param sheetNum     sheet number
	 * @param sheetName    name for the sheet
	 * @param cellDataList cell data list
	 * @return true/false
	 */
	public static boolean createExcel(final Workbook workbook, int sheetNum, String sheetName,
			List<CellData> cellDataList) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (sheetNum < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, sheetName);
			for (CellData cellData : cellDataList) {
				Common.insertCellValue(sheet, cellData.getColumn(), cellData.getRow(),
						ValueUtils.formatValue(cellData.getData()),
						Common.createCellStyle(workbook, cellData.getCellStyle().getCode()));
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
	 * @param cellDataList cell data list
	 * @return true/false
	 */
	public static boolean insertExcelData(final Workbook workbook, int sheetNum, List<CellData> cellDataList) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (sheetNum < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			for (CellData cellData : cellDataList) {
				Common.insertCellValue(sheet, cellData.getColumn(), cellData.getRow(),
						ValueUtils.formatValue(cellData.getData()),
						Common.createCellStyle(workbook, cellData.getCellStyle().getCode()));
			}
			flag = true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}
}
