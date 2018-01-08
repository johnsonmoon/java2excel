package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Writer
 * <p>
 * <pre>
 * 	Write excel file at multiple sheets.
 *
 * 	1. invoke writeExcelModel write model information. {@link Writer#writeExcelModel(Class, int)}
 * 	2. invoke writeExcelData append data into workbook. {@link Writer#writeExcelData(List, int)}
 * 	3. invoke flush to write workbook into disk file. {@link Writer#flush()} {@link Writer#flush(String)}
 * 	4. invoke close to close writer. {@link Writer#close}
 * </pre>
 * <p>
 * Created by xuyh at 2018/1/5 16:52.
 */
public class Writer extends AbstractWriter {
	private static Logger logger = LoggerFactory.getLogger(Writer.class);
	private Workbook workbook = new XSSFWorkbook();
	private Map<Integer, Integer> sheetCurrentRowNumberMap = new HashMap<>();
	private String filePathName;

	public Writer() {
	}

	public Writer(String filePathName) {
		this.filePathName = filePathName;
	}

	private Integer getSheetCurrentRowNumber(int sheetNumber) {
		if (!sheetCurrentRowNumberMap.containsKey(sheetNumber)) {
			sheetCurrentRowNumberMap.put(sheetNumber, 0);
			return 0;
		} else {
			return sheetCurrentRowNumberMap.get(sheetNumber);
		}
	}

	private void setSheetCurrentRowNumber(int sheetNumber, int currentRowNumber) {
		sheetCurrentRowNumberMap.put(sheetNumber, currentRowNumber);
	}

	/**
	 * Write model info into excel workbook at sheet number sheetNumber.
	 *
	 * @param clazz       given type
	 * @param sheetNumber given sheet number
	 * @return true/false
	 */
	public boolean writeExcelModel(Class<?> clazz, int sheetNumber) {
		if (workbook == null)
			return false;
		if (sheetNumber < 0)
			return false;
		return writeModel(clazz, workbook, sheetNumber);
	}

	/**
	 * Append data into excel workbook at sheet number sheetNumber.
	 *
	 * @param tList       data instance list
	 * @param sheetNumber given sheet number
	 * @return true/false
	 */
	public boolean writeExcelData(List<?> tList, int sheetNumber) {
		if (workbook == null)
			return false;
		if (sheetNumber < 0)
			return false;
		int currentRowNumber = getSheetCurrentRowNumber(sheetNumber);
		boolean flag = writeData(tList, workbook, sheetNumber, currentRowNumber);
		if (flag)
			setSheetCurrentRowNumber(sheetNumber, currentRowNumber + tList.size());
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
