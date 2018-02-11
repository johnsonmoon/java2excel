package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.formatted.model.Model;
import xuyihao.java2excel.util.FileUtils;

import java.util.ArrayList;
import java.util.List;

/**
 * Edit (Read & Write) excel file.
 * <p>
 * Created by xuyh at 2018/1/10 18:11.
 */
public class FormattedEditor extends FormattedAbstractEditor {
	private static Logger logger = LoggerFactory.getLogger(FormattedEditor.class);
	private Workbook workbook;
	private String filePathName;
	private String saveFilepathName;

	public FormattedEditor(String filePathName) {
		this.filePathName = filePathName;
		openWorkBook(filePathName);
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
	 * Read model info at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @return model instance (nullable)
	 */
	public Model readExcelModel(int sheetNumber) {
		Model model = null;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			model = readModel(workbook, sheetNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		} finally {
			close(workbook);
		}
		return model;
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
	 * Read model data at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @param beginRow    row to start read at
	 * @param readSize    given size to read
	 * @return model list of given size
	 */
	public List<Model> readExcelData(int sheetNumber, int beginRow, int readSize) {
		List<Model> models = new ArrayList<>();
		if (readSize <= 0)
			return models;
		if (sheetNumber < 0)
			return models;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			models.addAll(readData(workbook, sheetNumber, beginRow, readSize));
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return models;
	}

	/**
	 * Read type data at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @param beginRow    row to start read at
	 * @param readSize    given size to read
	 * @param clazz       given type
	 * @return type data list of given size
	 */
	public <T> List<T> readExcelData(int sheetNumber, int beginRow, int readSize, Class<T> clazz) {
		List<T> tList = new ArrayList<>();
		if (readSize <= 0)
			return tList;
		if (sheetNumber < 0)
			return tList;
		try {
			if (workbook == null)
				openWorkBook(filePathName);
			tList.addAll(readData(clazz, workbook, sheetNumber, beginRow, readSize));
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return tList;
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
	 * @param beginRow    row to start write at
	 * @return true/false
	 */
	public boolean writeExcelData(List<?> tList, int sheetNumber, int beginRow) {
		if (workbook == null)
			return false;
		if (sheetNumber < 0)
			return false;
		return writeData(tList, workbook, sheetNumber, beginRow);
	}

	/**
	 * Save changes.
	 *
	 * @return true/false
	 */
	public boolean flush() {
		this.saveFilepathName = this.filePathName;
		return flushOnly();
	}

	/**
	 * Save changes into another disk file.
	 *
	 * @param filePathName another given disk file path name (create new file)
	 * @return true/false
	 */
	public boolean flush(String filePathName) {
		this.saveFilepathName = filePathName;
		return flushOnly();
	}

	private boolean flushOnly() {
		if (workbook == null)
			return false;
		if (filePathName == null)
			return false;
		boolean flag = false;

		try {
			flag = flush(workbook, saveFilePathName());
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
		boolean flag;
		if (filePathName.equals(saveFilepathName)) {
			//save changes
			flag = close(workbook);
			flag = flag && FileUtils.delete(filePathName);
			flag = flag && FileUtils.rename(saveFilePathName(), filePathName);
			flag = flag && FileUtils.delete(backupFilePathName());
		} else {
			//save changes to another file
			flag = close(workbook);
			flag = flag && FileUtils.delete(filePathName);
			flag = flag && FileUtils.rename(saveFilePathName(), saveFilepathName);
			flag = flag && FileUtils.rename(backupFilePathName(), filePathName);
		}
		return flag;
	}

	private boolean openWorkBook(String filePathName) {
		try {
			//backup the old file
			FileUtils.copyFile(filePathName, backupFilePathName());

			if (workbook != null) {
				close(workbook);
			}
			if (FileUtils.exists(filePathName)) {
				workbook = new XSSFWorkbook(filePathName);
			} else {
				workbook = new XSSFWorkbook();
			}
			return true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
	}

	private String saveFilePathName() {
		return FileUtils.addHead2Name(saveFilepathName, "save_");
	}

	private String backupFilePathName() {
		return FileUtils.addHead2Name(filePathName, "backup_");
	}
}
