package com.github.johnsonmoon.java2excel.core;

import com.github.johnsonmoon.java2excel.core.entity.custom.ColumnMapper;
import com.github.johnsonmoon.java2excel.util.FileUtils;
import com.github.johnsonmoon.java2excel.util.ReflectionUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.github.johnsonmoon.java2excel.Editor;

import java.util.List;
import java.util.Map;

/**
 * Edit (Read & Write) excel file.
 * <p>
 * Created by xuyh at 2018/2/11 17:35.
 */
public class CustomEditor extends CustomAbstractEditor implements Editor {
	private static Logger logger = LoggerFactory.getLogger(CustomEditor.class);
	private Workbook workbook;
	private String filePathName;
	private String saveFilepathName;
	private int dataBeginRowNumber = 0;

	public CustomEditor(String filePathName) {
		this.filePathName = filePathName;
		openWorkBook(filePathName);
	}

	public CustomEditor(String filePathName, int dataBeginRowNumber) {
		this.filePathName = filePathName;
		this.dataBeginRowNumber = dataBeginRowNumber;
		openWorkBook(filePathName);
	}

	public int getDataBeginRowNumber() {
		return dataBeginRowNumber;
	}

	public void setDataBeginRowNumber(int dataBeginRowNumber) {
		this.dataBeginRowNumber = dataBeginRowNumber;
	}

	/**
	 * Write excel header by columnMapper information.
	 *
	 * @param columnMapper given column mapping information
	 * @param sheetNumber  given sheet number
	 * @param sheetName    given start row number to write with
	 * @return true/false
	 */
	public boolean writeExcelMetaInfo(ColumnMapper columnMapper, int sheetNumber, String sheetName) {
		if (columnMapper == null || workbook == null || sheetNumber < 0 || sheetName == null || sheetName.isEmpty()) {
			return false;
		}
		return writeColumnMapperHeader(columnMapper, workbook, sheetNumber, sheetName);
	}

	/**
	 * Write excel header by columnMapper information.
	 *
	 * @param clazz       type
	 * @param sheetNumber given sheet number
	 * @param sheetName   given start row number to write with
	 * @return true/false
	 */
	public boolean writeExcelMetaInfo(Class<?> clazz, int sheetNumber,
			String sheetName) {
		if (clazz == null || workbook == null || sheetNumber < 0 || sheetName == null || sheetName.isEmpty()) {
			return false;
		}
		return writeColumnMapperHeader(clazz, workbook, sheetNumber, sheetName);
	}

	/**
	 * Write excel header by columnMapper information.
	 *
	 * @param clazz       type
	 * @param sheetNumber given sheet number
	 * @return true/false
	 */
	@Override
	public boolean writeExcelMetaInfo(Class<?> clazz, int sheetNumber) {
		if (clazz == null || workbook == null || sheetNumber < 0) {
			return false;
		}
		return writeColumnMapperHeader(clazz, workbook, sheetNumber, ReflectionUtils.getClassNameShort(clazz));
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param dataMapList    given data map list
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	public boolean writeExcelDataDirectly(List<Map<Integer, Object>> dataMapList, int sheetNumber,
			int startRowNumber) {
		if (dataMapList == null || dataMapList.isEmpty() || workbook == null || sheetNumber < 0 || startRowNumber < 0)
			return false;
		boolean flag = false;
		try {
			flag = writeDataDirectly(dataMapList, workbook, sheetNumber, startRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return flag;
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param dataMapList    given data map list
	 * @param columnMapper   given column mapping information
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	public boolean writeExcelDataDirectly(List<Map<String, Object>> dataMapList, ColumnMapper columnMapper,
			int sheetNumber, int startRowNumber) {
		if (dataMapList == null || dataMapList.isEmpty() || columnMapper == null || workbook == null || sheetNumber < 0
				|| startRowNumber < 0) {
			return false;
		}
		boolean flag = false;
		try {
			flag = writeDataDirectly(dataMapList, columnMapper, workbook, sheetNumber, startRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return flag;
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param tList          given clazz data list
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	@Override
	public boolean writeExcelData(List<?> tList, int sheetNumber, int startRowNumber) {
		if (tList == null || tList.isEmpty() || workbook == null || sheetNumber < 0 || startRowNumber < 0) {
			return false;
		}
		boolean flag = false;
		try {
			flag = writeData(tList, workbook, sheetNumber, startRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return flag;
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param tList          given clazz data list
	 * @param columnMapper   column mapping information
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	public boolean writeExcelData(List<?> tList, ColumnMapper columnMapper, int sheetNumber,
			int startRowNumber) {
		if (tList == null || tList.isEmpty() || columnMapper == null || workbook == null || sheetNumber < 0
				|| startRowNumber < 0) {
			return false;
		}
		boolean flag = false;
		try {
			flag = writeData(tList, columnMapper, workbook, sheetNumber, startRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return flag;
	}

	/**
	 * Read map data from a given workbook in sheet from a sheetNumber.
	 *
	 * @param sheetNumber    given sheet number
	 * @param beginRowNumber row number to begin reading
	 * @param readSize       row count to read
	 * @return map list
	 */
	public List<Map<Integer, Object>> readExcelDataDirectly(int sheetNumber, int beginRowNumber,
			int readSize) {
		if (workbook == null || sheetNumber < 0 || beginRowNumber < 0)
			return null;
		if (readSize < 0)
			throw new RuntimeException("readSize can not below 0.");
		List<Map<Integer, Object>> mapList = null;
		try {
			mapList = readDataDirectly(workbook, sheetNumber, beginRowNumber, readSize);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return mapList;
	}

	/**
	 * Read map data from a given workbook in sheet from a sheetNumber.
	 *
	 * @param columnMapper   column mapping information
	 * @param sheetNumber    given sheet number
	 * @param beginRowNumber row number to begin reading
	 * @param readSize       row count to read
	 * @return map list
	 */
	public List<Map<String, Object>> readExcelDataDirectly(ColumnMapper columnMapper,
			int sheetNumber,
			int beginRowNumber, int readSize) {
		if (columnMapper == null || workbook == null || sheetNumber < 0 || beginRowNumber < 0 || readSize < 0)
			return null;
		List<Map<String, Object>> mapList = null;
		try {
			mapList = readDataDirectly(columnMapper, workbook, sheetNumber, beginRowNumber, readSize);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return mapList;
	}

	/**
	 * Read type data from a given workbook in sheet from a sheetNumber.
	 *
	 * @param clazz          type
	 * @param columnMapper   column mapping information
	 * @param sheetNumber    given sheet number
	 * @param beginRowNumber row number to begin reading
	 * @param readSize       row count to read
	 * @param <T>            type
	 * @return type list
	 */
	public <T> List<T> readExcelData(Class<T> clazz, ColumnMapper columnMapper, int sheetNumber,
			int beginRowNumber, int readSize) {
		if (clazz == null || columnMapper == null || workbook == null || sheetNumber < 0 || beginRowNumber < 0
				|| readSize <= 0)
			return null;
		try {
			return readData(clazz, columnMapper, workbook, sheetNumber, beginRowNumber, readSize);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return null;
	}

	/**
	 * Read type data from a given workbook in sheet from a sheetNumber.
	 *
	 * @param clazz          type
	 * @param sheetNumber    given sheet number
	 * @param beginRowNumber row number to begin reading
	 * @param readSize       row count to read
	 * @param <T>            type
	 * @return type list
	 */
	@Override
	public <T> List<T> readExcelData(int sheetNumber, int beginRowNumber,
			int readSize, Class<T> clazz) {
		if (clazz == null || workbook == null || sheetNumber < 0 || beginRowNumber < 0 || readSize <= 0)
			return null;
		try {
			return readData(clazz, workbook, sheetNumber, beginRowNumber, readSize);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return null;
	}

	/**
	 * Read data count from a given workbook in sheet from a sheetNumber where data begins at row.
	 *
	 * @param sheetNumber        given sheet number
	 * @param dataBeginRowNumber row number which data begins
	 * @return data count
	 */
	public int readExcelDataCount(int sheetNumber, int dataBeginRowNumber) {
		if (workbook == null || sheetNumber < 0 || dataBeginRowNumber < 0)
			return 0;
		try {
			return readDataCount(workbook, sheetNumber, dataBeginRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return 0;
	}

	/**
	 * Read data count from a given workbook at sheet.
	 *
	 * @param sheetNumber given sheet number
	 * @return data count
	 */
	@Override
	public int readExcelDataCount(int sheetNumber) {
		if (workbook == null || sheetNumber < 0)
			return 0;
		try {
			return readDataCount(workbook, sheetNumber, dataBeginRowNumber);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
		return 0;
	}

	/**
	 * Save changes.
	 *
	 * @return true/false
	 */
	@Override
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
	@Override
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
	@Override
	public void close() {
		if (workbook == null)
			return;
		if (saveFilepathName == null) {
			close(workbook);
		} else if (filePathName.equals(saveFilepathName)) {
			//save changes
			close(workbook);
			FileUtils.delete(filePathName);
			FileUtils.rename(saveFilePathName(), filePathName);
			FileUtils.delete(backupFilePathName());
		} else {
			//save changes to another file
			if (FileUtils.exists(saveFilepathName))
				FileUtils.delete(saveFilepathName);
			close(workbook);
			FileUtils.delete(filePathName);
			FileUtils.rename(saveFilePathName(), saveFilepathName);
			FileUtils.rename(backupFilePathName(), filePathName);
		}
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
		return FileUtils.addHead2Name(saveFilepathName, "custom_save_");
	}

	private String backupFilePathName() {
		return FileUtils.addHead2Name(filePathName, "custom_backup_");
	}
}
