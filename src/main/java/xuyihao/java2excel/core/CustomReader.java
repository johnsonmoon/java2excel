package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Custom reader.
 * <p>
 * Created by xuyh at 2018/2/11 17:35.
 */
public class CustomReader extends CustomAbstractReader {
	private static Logger logger = LoggerFactory.getLogger(CustomReader.class);
	private Workbook workbook;
	private String filePathName;
	private int dataBeginRowNumber = 0;
	private Map<Integer, Integer> typeSheetCurrentRowNumberMap = new HashMap<>();

	public CustomReader(String filePathName, int dataBeginRowNumber) {
		this.filePathName = filePathName;
		this.dataBeginRowNumber = dataBeginRowNumber;
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

	private void setTypeSheetCurrentRowNumber(int sheetNumber, int currentRowNumber) {
		typeSheetCurrentRowNumberMap.put(sheetNumber, currentRowNumber);
	}

	private Integer getTypeSheetCurrentRowNumber(int sheetNumber) {
		if (!typeSheetCurrentRowNumberMap.containsKey(sheetNumber)) {
			typeSheetCurrentRowNumberMap.put(sheetNumber, dataBeginRowNumber);
			return dataBeginRowNumber;
		} else {
			return typeSheetCurrentRowNumberMap.get(sheetNumber);
		}
	}

	public int getDataBeginRowNumber() {
		return dataBeginRowNumber;
	}

	public void setDataBeginRowNumber(int dataBeginRowNumber) {
		this.dataBeginRowNumber = dataBeginRowNumber;
	}

	/**
	 * Get data count at sheet sheetNumber
	 *
	 * @param sheetNumber sheet number
	 * @return data count
	 */
	public int readExcelDataCount(int sheetNumber) {
		int dataCount = 0;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			dataCount = readDataCount(workbook, sheetNumber, dataBeginRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return dataCount;
	}

	/**
	 * Read type data at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @param readSize    given size to read
	 * @param clazz       given type
	 * @return type data list of given size
	 */
	public <T> List<T> readExcelData(int sheetNumber, int readSize, Class<T> clazz) {
		List<T> tList = new ArrayList<>();
		if (readSize <= 0)
			return tList;
		if (sheetNumber < 0)
			return tList;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			int currentRowNumber = getTypeSheetCurrentRowNumber(sheetNumber);
			tList.addAll(readData(clazz, workbook, sheetNumber, currentRowNumber, readSize));
			setTypeSheetCurrentRowNumber(sheetNumber, currentRowNumber + tList.size());
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return tList;
	}

	/**
	 * Read type data at sheet sheetNumber and filling data into given type data array.
	 *
	 * @param sheetNumber sheet number
	 * @param ts          given model array [not null]
	 * @param clazz       given type
	 * @return read size
	 */
	public <T> int readExcelData(int sheetNumber, T[] ts, Class<T> clazz) {
		int readCount = 0;
		if (ts == null)
			return readCount;
		if (ts.length == 0)
			return readCount;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			int arrayLength = ts.length;
			int currentRowNumber = getTypeSheetCurrentRowNumber(sheetNumber);
			List<T> tList = readData(clazz, workbook, sheetNumber, currentRowNumber, arrayLength);
			if (tList == null)
				return readCount;
			readCount = tList.size();
			if (readCount != 0)
				setTypeSheetCurrentRowNumber(sheetNumber, currentRowNumber + readCount);
			for (int i = 0; i < readCount; i++) {
				ts[i] = tList.get(i);
			}
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return readCount;
	}

	/**
	 * Refresh reader index of the excel file sheet position (row number).
	 * <p>
	 * <pre>
	 *     Using for re-read the excel file (all sheet) from the begin row. {@link CustomReader#dataBeginRowNumber}
	 * </pre>
	 *
	 * @return true/false
	 */
	public boolean refresh() {
		try {
			for (Integer integer : typeSheetCurrentRowNumberMap.keySet()) {
				setTypeSheetCurrentRowNumber(integer, dataBeginRowNumber);
			}
			return true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
	}

	/**
	 * Close the workbook.
	 *
	 * @return true/false
	 */
	public boolean close() {
		if (workbook == null)
			return false;
		return close(workbook);
	}
}
