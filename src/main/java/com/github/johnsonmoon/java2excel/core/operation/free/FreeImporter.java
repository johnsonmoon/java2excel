package com.github.johnsonmoon.java2excel.core.operation.free;

import com.github.johnsonmoon.java2excel.core.entity.free.CellData;
import com.github.johnsonmoon.java2excel.core.operation.Common;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created by xuyh at 2018/3/20 15:02.
 */
public class FreeImporter {
	/**
	 * Get cell data from workbook at sheet number sheetNumber
	 *
	 * @param workbook    excel workbook
	 * @param sheetNumber sheet number
	 * @param row         row number
	 * @param column      column number
	 * @return cellData instance
	 */
	public static CellData getCellDataFromExcel(final Workbook workbook, int sheetNumber, int row, int column) {
		if (workbook == null)
			return null;
		if (sheetNumber < 0)
			return null;
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		if (sheet == null)
			return null;
		String str = Common.getCellValue(sheet, column, row);
		return new CellData(row, column, str, null);
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
	public static String getDataFromExcel(final Workbook workbook, int sheetNumber, int row, int column) {
		if (workbook == null)
			return null;
		if (sheetNumber < 0)
			return null;
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		if (sheet == null)
			return null;
		return Common.getCellValue(sheet, column, row);
	}
}
