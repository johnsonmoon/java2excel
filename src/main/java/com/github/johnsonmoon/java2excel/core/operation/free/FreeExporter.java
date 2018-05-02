package com.github.johnsonmoon.java2excel.core.operation.free;

import com.github.johnsonmoon.java2excel.core.entity.free.CellData;
import com.github.johnsonmoon.java2excel.core.operation.Common;
import com.github.johnsonmoon.java2excel.util.ValueUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * Created by xuyh at 2018/3/20 15:02.
 */
public class FreeExporter {
	private static Logger logger = LoggerFactory.getLogger(FreeExporter.class);

	/**
	 * Merge cells begin from {firstRow}, {lastRow} to {firstColumn}, {lastColumn}.
	 *
	 * @param workbook    given workbook
	 * @param sheetNum    given sheet number
	 * @param firstRow    first row index
	 * @param lastRow     last row index
	 * @param firstColumn first column index
	 * @param lastColumn  last column index
	 * @return true/false
	 */
	public static boolean mergeCells(final Workbook workbook, int sheetNum, int firstRow, int lastRow, int firstColumn,
			int lastColumn) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (sheetNum < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			flag = Common.mergeCells(sheet, firstRow, lastRow, firstColumn, lastColumn);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Set default row height of the sheet.
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param height   height
	 * @return true/false
	 */
	public static boolean setDefaultRowHeight(final Workbook workbook, int sheetNum, int height) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (sheetNum < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			flag = Common.setDefaultRowHeight(sheet, height);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Set height of row. Row should be exists.
	 * <pre>
	 *     NOTE:invoke after row is created.
	 * </pre>
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param row      row number, begin from 0.
	 * @param height   height
	 * @return true/false
	 */
	public static boolean setRowHeight(final Workbook workbook, int sheetNum, int row, int height) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (sheetNum < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			flag = Common.setRowHeight(sheet, row, height);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Hide row, Set height 0 of row. Row should be exists.
	 * <pre>
	 *     NOTE:invoke after row is created.
	 * </pre>
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param row      row number, begin from 0.
	 * @return true/false
	 */
	public static boolean hideRow(final Workbook workbook, int sheetNum, int row) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (sheetNum < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			flag = Common.hideRow(sheet, row);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Set width of column.
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param column   column number, begin from 0.
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
			flag = Common.setColumnWidth(sheet, column, width);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Hide column, Set width 0 of column.
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param column   column number, begin from 0.
	 * @return true/false
	 */
	public static boolean hideColumn(final Workbook workbook, int sheetNum, int column) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (column < 0)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			flag = Common.hideColumn(sheet, column);
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
