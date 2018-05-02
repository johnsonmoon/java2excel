package com.github.johnsonmoon.java2excel;

import java.util.List;

/**
 * Created by xuyh at 2018/2/12 16:28.
 */
public interface Editor {
	/**
	 * Write excel header meta information by type clazz info.
	 *
	 * @param clazz       type
	 * @param sheetNumber given sheet number
	 * @return true/false
	 */
	boolean writeExcelMetaInfo(Class<?> clazz, int sheetNumber);

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param tList          given clazz data list
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	boolean writeExcelData(List<?> tList, int sheetNumber, int beginRow);

	/**
	 * Read data count from a given workbook at sheet.
	 *
	 * @param sheetNumber given sheet number
	 * @return data count
	 */
	int readExcelDataCount(int sheetNumber);

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
	<T> List<T> readExcelData(int sheetNumber, int beginRowNumber,
			int readSize, Class<T> clazz);

	/**
	 * Save changes.
	 *
	 * @return true/false
	 */
	boolean flush();

	/**
	 * Save changes into another disk file.
	 *
	 * @param filePathName another given disk file path name (create new file)
	 * @return true/false
	 */
	boolean flush(String filePathName);

	/**
	 * Close the writer
	 *
	 * @return true/false
	 */
	boolean close();
}
