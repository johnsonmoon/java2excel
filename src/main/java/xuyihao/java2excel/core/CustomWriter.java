package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.util.ReflectionUtils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/2/11 17:35.
 */
public class CustomWriter extends CustomAbstractWriter {
	private static Logger logger = LoggerFactory.getLogger(CustomWriter.class);
	private Workbook workbook = new XSSFWorkbook();
	private Map<Integer, Integer> sheetCurrentRowNumberMap = new HashMap<>();
	private String filePathName;

	public CustomWriter() {
	}

	public CustomWriter(String filePathName) {
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
	public boolean writeExcelMetaInfo(Class<?> clazz, int sheetNumber) {
		if (workbook == null)
			return false;
		if (sheetNumber < 0)
			return false;
		boolean flag = writeColumnMapperHeader(clazz, workbook, sheetNumber, ReflectionUtils.getClassNameShort(clazz));
		if (flag)
			setSheetCurrentRowNumber(sheetNumber, 1);
		return flag;
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
