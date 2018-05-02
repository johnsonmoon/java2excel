package com.github.johnsonmoon.java2excel.core;

import com.github.johnsonmoon.java2excel.core.entity.free.CellData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Closeable;

/**
 * Created by xuyh at 2018/3/20 16:46.
 */
public class FreeReader extends FreeAbstractReader implements Closeable {
	private static Logger logger = LoggerFactory.getLogger(FreeReader.class);
	private Workbook workbook;

	public FreeReader(String filePathName) {
		openWorkBook(filePathName);
	}

	private boolean openWorkBook(String filePathName) {
		try {
			if (workbook != null) {
				close(workbook);
			}
			workbook = new XSSFWorkbook(filePathName);
			return true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
	}

	/**
	 * Get cell data from workbook at sheet number sheetNumber
	 *
	 * @param sheetNumber sheet number
	 * @param row         row number
	 * @param column      column number
	 * @return cellData instance
	 */
	public CellData readExcelCellData(int sheetNumber, int row, int column) {
		if (workbook == null || sheetNumber < 0) {
			return null;
		}
		try {
			return readCellData(workbook, sheetNumber, row, column);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return null;
		}
	}

	/**
	 * Get string data from workbook at sheet number sheetNumber
	 *
	 * @param sheetNumber sheet number
	 * @param row         row number
	 * @param column      column number
	 * @return string data
	 */
	public String readExcelData(int sheetNumber, int row, int column) {
		if (workbook == null || sheetNumber < 0) {
			return null;
		}
		try {
			return readData(workbook, sheetNumber, row, column);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return null;
		}
	}

	/**
	 * Close the workbook.
	 *
	 * @return true/false
	 */
	public void close() {
		if (workbook == null)
			return;
		close(workbook);
	}
}
