package com.github.johnsonmoon.java2excel;

import java.io.Closeable;
import java.util.List;

/**
 * Excel writer.
 * <p>
 * <p>
 * <pre>
 * 	Write excel file at multiple sheets as a whole new excel file.
 *
 * 	1. invoke writeExcelMetaInfo write model information. {@link Writer#writeExcelMetaInfo(Class, int)}
 * 	2. invoke writeExcelData append data into workbook. {@link Writer#writeExcelData(List, int)}
 * 	3. invoke flush to write workbook into disk file. {@link Writer#flush()} {@link Writer#flush(String)}
 * 	4. invoke close to close writer. {@link Writer#close}
 * </pre>
 * <p>
 * Created by xuyh at 2018/2/12 16:07.
 */
public interface Writer extends Closeable {
	/**
	 * Write model info into excel workbook at sheet number sheetNumber.
	 *
	 * @param clazz       given type
	 * @param sheetNumber given sheet number
	 * @return true/false
	 */
	boolean writeExcelMetaInfo(Class<?> clazz, int sheetNumber);

	/**
	 * Append data into excel workbook at sheet number sheetNumber.
	 *
	 * @param tList       data instance list
	 * @param sheetNumber given sheet number
	 * @return true/false
	 */
	boolean writeExcelData(List<?> tList, int sheetNumber);

	/**
	 * Write workbook into given disk file.
	 *
	 * @return true/false
	 */
	boolean flush();

	/**
	 * Write workbook into given disk file.
	 *
	 * @param filePathName given disk file path name (create new file)
	 * @return true/false
	 */
	boolean flush(String filePathName);
}
