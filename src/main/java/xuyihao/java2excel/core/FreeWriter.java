package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.free.CellData;

import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/3/20 16:46.
 */
public class FreeWriter extends FreeAbstractWriter {
	private static Logger logger = LoggerFactory.getLogger(FreeWriter.class);
	private Workbook workbook = new XSSFWorkbook();
	private String filePathName;

	public FreeWriter() {
	}

	public FreeWriter(String filePathName) {
		this.filePathName = filePathName;
	}

	public void setFilePathName(String filePathName) {
		this.filePathName = filePathName;
	}

	/**
	 * Set column width for workbook at sheetNum column.
	 *
	 * @param sheetNumber given sheet number
	 * @param column      given column number
	 * @param columnWidth width
	 * @return true/false
	 */
	public boolean setExcelColumnWidth(int sheetNumber, int column, int columnWidth) {
		if (workbook == null || sheetNumber < 0) {
			return false;
		}
		return setColumnWidth(workbook, sheetNumber, column, columnWidth);
	}

	/**
	 * Create a excel sheet and write data into a workbook with given cellData list.
	 *
	 * @param cellDataList cell data list
	 * @param sheetNumber  given sheet number
	 * @param sheetName    given sheet name
	 * @return true/false
	 */
	public boolean createExcel(List<CellData> cellDataList, int sheetNumber, String sheetName) {
		if (workbook == null || sheetNumber < 0) {
			return false;
		}
		boolean flag;
		try {
			flag = create(cellDataList, workbook, sheetNumber, sheetName);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Write data into a workbook with given cellData list.
	 *
	 * @param cellDataList cell data list
	 * @param sheetNumber  given sheet number
	 * @return true/false
	 */
	public boolean writeExcelData(List<CellData> cellDataList, int sheetNumber) {
		if (workbook == null || sheetNumber < 0) {
			return false;
		}
		boolean flag;
		try {
			flag = writeData(cellDataList, workbook, sheetNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Write multiple row data into a workbook's given sheet starting from given start row number.
	 *
	 * @param dataMapList    given data map list
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	public boolean writeExcelDataDirectly(List<Map<Integer, Object>> dataMapList, int sheetNumber, int startRowNumber) {
		if (workbook == null || sheetNumber < 0) {
			return false;
		}
		boolean flag;
		try {
			flag = writeDataDirectly(dataMapList, workbook, sheetNumber, startRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Write single row data into a workbook's given sheet at a given row number.
	 *
	 * @param dataMap     given data map (for one row)
	 * @param sheetNumber given sheet number
	 * @param rowNumber   given row number to write with
	 * @return true/false
	 */
	public boolean writeExcelDataDirectly(Map<Integer, Object> dataMap, int sheetNumber,
			int rowNumber) {
		if (workbook == null || sheetNumber < 0) {
			return false;
		}
		boolean flag;
		try {
			flag = writeDataDirectly(dataMap, workbook, sheetNumber, rowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * Write workbook into given disk file.
	 *
	 * @return true/false
	 */
	public boolean flush() {
		return flushOnly();
	}

	/**
	 * Write workbook into given disk file.
	 *
	 * @param filePathName given disk file path name (create new file)
	 * @return true/false
	 */
	public boolean flush(String filePathName) {
		this.filePathName = filePathName;
		return flushOnly();
	}

	private boolean flushOnly() {
		if (workbook == null)
			return false;
		if (filePathName == null)
			return false;
		boolean flag = false;
		try {
			flag = flush(workbook, filePathName);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return flag;
	}

	/**
	 * Close the writer
	 *
	 * @return true/false
	 */
	public boolean close() {
		if (workbook == null)
			return false;
		return close(workbook);
	}
}
