package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.model.Template;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Reader
 * <p>
 * <pre>
 * 	Read excel file at multiple sheets.
 * </pre>
 * <p>
 * Created by xuyh at 2018/1/5 16:52.
 */
public class Reader extends AbstractReader {
	private static Logger logger = LoggerFactory.getLogger(Reader.class);
	private Workbook workbook;
	private String filePathName;
	private Map<Integer, Integer> templateSheetCurrentRowNumberMap = new HashMap<>();
	private Map<Integer, Integer> typeSheetCurrentRowNumberMap = new HashMap<>();

	public Reader(String filePathName) {
		this.filePathName = filePathName;
		openWorkBook(filePathName);
	}

	private void setTemplateSheetCurrentRowNumber(int sheetNumber, int currentRowNumber) {
		templateSheetCurrentRowNumberMap.put(sheetNumber, currentRowNumber);
	}

	private Integer getTemplateSheetCurrentRowNumber(int sheetNumber) {
		if (!templateSheetCurrentRowNumberMap.containsKey(sheetNumber)) {
			templateSheetCurrentRowNumberMap.put(sheetNumber, 0);
			return 0;
		} else {
			return templateSheetCurrentRowNumberMap.get(sheetNumber);
		}
	}

	private void setTyepSheetCurrentRowNumber(int sheetNumber, int currentRowNumber) {
		typeSheetCurrentRowNumberMap.put(sheetNumber, currentRowNumber);
	}

	private Integer getTypeSheetCurrentRowNumber(int sheetNumber) {
		if (!typeSheetCurrentRowNumberMap.containsKey(sheetNumber)) {
			typeSheetCurrentRowNumberMap.put(sheetNumber, 0);
			return 0;
		} else {
			return typeSheetCurrentRowNumberMap.get(sheetNumber);
		}
	}

	/**
	 * Read type info at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @return type (nullable)
	 */
	public Class<?> readExcelJavaClass(int sheetNumber) {
		Class<?> clazz = null;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			clazz = readJavaClass(workbook, sheetNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		} finally {
			close(workbook);
		}
		return clazz;
	}

	/**
	 * Read template info at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @return template (nullable)
	 */
	public Template readExcelTemplate(int sheetNumber) {
		Template template = null;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			template = readTemplate(workbook, sheetNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		} finally {
			close(workbook);
		}
		return template;
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
			dataCount = readDataCount(workbook, sheetNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return dataCount;
	}

	/**
	 * Read template data at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @param readSize    given size to read
	 * @return template list of given size
	 */
	public List<Template> readExcelData(int sheetNumber, int readSize) {
		List<Template> templates = new ArrayList<>();
		if (readSize <= 0)
			return templates;
		if (sheetNumber < 0)
			return templates;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			int currentRowNumber = getTemplateSheetCurrentRowNumber(sheetNumber);
			templates.addAll(readData(workbook, sheetNumber, currentRowNumber, readSize));
			setTemplateSheetCurrentRowNumber(sheetNumber, currentRowNumber + templates.size());
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return templates;
	}

	/**
	 * Read template data at sheet sheetNumber and filling data into given template array.
	 *
	 * @param sheetNumber sheet number
	 * @param templates   given template array [not null]
	 * @return read size
	 */
	public int readExcelData(int sheetNumber, Template[] templates) {
		int readCount = 0;
		if (templates == null)
			return readCount;
		if (templates.length == 0)
			return readCount;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			int arrayLength = templates.length;
			int currentRowNumber = getTemplateSheetCurrentRowNumber(sheetNumber);
			List<Template> templateList = readData(workbook, sheetNumber, currentRowNumber, arrayLength);
			if (templateList == null)
				return readCount;
			readCount = templateList.size();
			if (readCount != 0)
				setTemplateSheetCurrentRowNumber(sheetNumber, currentRowNumber + readCount);
			for (int i = 0; i < readCount; i++) {
				templates[i] = templateList.get(i);
			}
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return readCount;
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
			setTyepSheetCurrentRowNumber(sheetNumber, currentRowNumber + tList.size());
		} catch (Exception e) {
			logger.warn(e.getMessage());
		}
		return tList;
	}

	/**
	 * Read type data at sheet sheetNumber and filling data into given type data array.
	 *
	 * @param sheetNumber sheet number
	 * @param ts          given template array [not null]
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
				setTyepSheetCurrentRowNumber(sheetNumber, currentRowNumber + readCount);
			for (int i = 0; i < readCount; i++) {
				ts[i] = tList.get(i);
			}
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return readCount;
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
	 * Refresh reader index of the excel file sheet position (row number).
	 * <p>
	 * <pre>
	 *     Using for re-read the excel file (all sheet) from the begin row.
	 * </pre>
	 *
	 * @return true/false
	 */
	public boolean refresh() {
		try {
			for (Integer integer : templateSheetCurrentRowNumberMap.keySet()) {
				setTemplateSheetCurrentRowNumber(integer, 0);
			}

			for (Integer integer : typeSheetCurrentRowNumberMap.keySet()) {
				setTyepSheetCurrentRowNumber(integer, 0);
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
