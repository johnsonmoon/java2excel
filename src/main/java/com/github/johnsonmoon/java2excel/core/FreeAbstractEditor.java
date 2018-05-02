package com.github.johnsonmoon.java2excel.core;

import com.github.johnsonmoon.java2excel.core.entity.common.CellStyle;
import com.github.johnsonmoon.java2excel.core.entity.free.CellData;
import com.github.johnsonmoon.java2excel.core.operation.Common;
import com.github.johnsonmoon.java2excel.core.operation.free.FreeExporter;
import com.github.johnsonmoon.java2excel.core.operation.free.FreeImporter;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.johnsonmoon.java2excel.util.FileUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/3/20 16:09.
 */
public abstract class FreeAbstractEditor {
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
	protected boolean mergeCells(final Workbook workbook, int sheetNum, int firstRow, int lastRow, int firstColumn,
			int lastColumn) {
		if (workbook == null)
			return false;
		if (sheetNum < 0)
			return false;
		return FreeExporter.mergeCells(workbook, sheetNum, firstRow, lastRow, firstColumn, lastColumn);
	}

	/**
	 * Set default row height of the sheet.
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param height   height
	 * @return true/false
	 */
	protected boolean setDefaultRowHeight(final Workbook workbook, int sheetNum, int height) {
		if (workbook == null)
			return false;
		if (sheetNum < 0)
			return false;
		return FreeExporter.setDefaultRowHeight(workbook, sheetNum, height);
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
	protected boolean setRowHeight(final Workbook workbook, int sheetNum, int row, int height) {
		if (workbook == null)
			return false;
		if (sheetNum < 0)
			return false;
		return FreeExporter.setRowHeight(workbook, sheetNum, row, height);
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
	protected boolean hideRow(final Workbook workbook, int sheetNum, int row) {
		if (workbook == null)
			return false;
		if (sheetNum < 0)
			return false;
		return FreeExporter.hideRow(workbook, sheetNum, row);
	}

	/**
	 * Set column width for workbook at sheetNum column.
	 *
	 * @param workbook    given workbook
	 * @param sheetNumber given sheet number
	 * @param column      given column number
	 * @param columnWidth width
	 * @return true/false
	 */
	protected boolean setColumnWidth(final Workbook workbook, int sheetNumber, int column, int columnWidth) {
		return FreeExporter.setColumnWidth(workbook, sheetNumber, column, columnWidth);
	}

	/**
	 * Hide column, Set width 0 of column.
	 *
	 * @param workbook given workbook
	 * @param sheetNum given sheet number
	 * @param column   column number, begin from 0.
	 * @return true/false
	 */
	protected boolean hideColumn(final Workbook workbook, int sheetNum, int column) {
		if (workbook == null)
			return false;
		if (sheetNum < 0)
			return false;
		return FreeExporter.hideColumn(workbook, sheetNum, column);
	}

	/**
	 * Create a excel sheet and write data into a workbook with given cellData list.
	 *
	 * @param cellDataList cell data list
	 * @param workbook     given workbook
	 * @param sheetNumber  given sheet number
	 * @param sheetName    given sheet name
	 * @return true/false
	 */
	protected boolean create(List<CellData> cellDataList, Workbook workbook, int sheetNumber, String sheetName)
			throws Exception {
		return FreeExporter.createExcel(workbook, sheetNumber, sheetName, cellDataList);
	}

	/**
	 * Write data into a workbook with given cellData list.
	 *
	 * @param cellDataList cell data list
	 * @param workbook     given workbook
	 * @param sheetNumber  given sheet number
	 * @return true/false
	 */
	protected boolean writeData(List<CellData> cellDataList, Workbook workbook, int sheetNumber) throws Exception {
		return FreeExporter.insertExcelData(workbook, sheetNumber, cellDataList);
	}

	/**
	 * Write multiple row data into a workbook's given sheet starting from given start row number.
	 *
	 * @param dataMapList    given data map list
	 * @param workbook       given workbook
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	protected boolean writeDataDirectly(List<Map<Integer, Object>> dataMapList, Workbook workbook, int sheetNumber,
			int startRowNumber) throws Exception {
		if (dataMapList == null || dataMapList.isEmpty())
			throw new RuntimeException("dataMapList is empty or null.");
		if (workbook == null)
			throw new NullPointerException("workbook is null.");
		if (sheetNumber < 0 || startRowNumber < 0)
			throw new RuntimeException("shettNumber or startRowNumber can not be less than 0.");
		List<CellData> cellDataList = new ArrayList<>();
		for (Map<Integer, Object> map : dataMapList) {
			for (Map.Entry<Integer, Object> entry : map.entrySet()) {
				cellDataList
						.add(new CellData(startRowNumber, entry.getKey(), entry.getValue(), CellStyle.CELL_STYLE_TYPE_VALUE));
			}
			startRowNumber++;
		}
		return FreeExporter.insertExcelData(workbook, sheetNumber, cellDataList);
	}

	/**
	 * Write single row data into a workbook's given sheet at a given row number.
	 *
	 * @param dataMap     given data map (for one row)
	 * @param workbook    given workbook
	 * @param sheetNumber given sheet number
	 * @param rowNumber   given row number to write with
	 * @return true/false
	 */
	protected boolean writeDataDirectly(Map<Integer, Object> dataMap, Workbook workbook, int sheetNumber,
			int rowNumber) throws Exception {
		if (dataMap == null || dataMap.isEmpty())
			throw new RuntimeException("dataMapList is empty or null.");
		if (workbook == null)
			throw new NullPointerException("workbook is null.");
		if (sheetNumber < 0 || rowNumber < 0)
			throw new RuntimeException("shettNumber or startRowNumber can not be less than 0.");
		List<CellData> cellDataList = new ArrayList<>();
		for (Map.Entry<Integer, Object> entry : dataMap.entrySet()) {
			cellDataList.add(new CellData(rowNumber, entry.getKey(), entry.getValue(), CellStyle.CELL_STYLE_TYPE_VALUE));
		}
		return FreeExporter.insertExcelData(workbook, sheetNumber, cellDataList);
	}

	/**
	 * Write existing workbook into a given file. Create a new file if the file isn't exist.
	 *
	 * @param workbook     given workbook
	 * @param filePathName given file path name
	 * @return true/false
	 * @throws Exception exceptions
	 */
	protected boolean flush(Workbook workbook, String filePathName) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook is null!");
		File file = FileUtils.createFile(filePathName);
		return Common.writeFileToDisk(workbook, file);
	}

	/**
	 * Get cell data from workbook at sheet number sheetNumber
	 *
	 * @param workbook    excel workbook
	 * @param sheetNumber sheet number
	 * @param row         row number
	 * @param column      column number
	 * @return cellData instance
	 */
	protected CellData readCellData(Workbook workbook, int sheetNumber, int row, int column) throws Exception {
		return FreeImporter.getCellDataFromExcel(workbook, sheetNumber, row, column);
	}

	/**
	 * Get string data from workbook at sheet number sheetNumber
	 *
	 * @param workbook    excel workbook
	 * @param sheetNumber sheet number
	 * @param row         row number
	 * @param column      column number
	 * @return string data
	 */
	protected String readData(Workbook workbook, int sheetNumber, int row, int column) throws Exception {
		return FreeImporter.getDataFromExcel(workbook, sheetNumber, row, column);
	}

	/**
	 * Close workbook, release resources.
	 *
	 * @param workbook given workbook
	 * @return true/false
	 */
	protected boolean close(Workbook workbook) {
		return Common.closeWorkbook(workbook);
	}
}
