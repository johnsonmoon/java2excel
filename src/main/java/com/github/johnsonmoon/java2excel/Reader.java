package com.github.johnsonmoon.java2excel;

import java.io.Closeable;
import java.util.List;

/**
 * Excel reader
 * <p>
 * <pre>
 * 	Read excel file at multiple sheets from a existing excel file.
 *
 * 	1.invoke readExcelDataCount to read data count from excel file. {@link Reader#readExcelDataCount(int)}
 * 	2.invoke readExcelData to read data from excel file. {@link Reader#readExcelData(int, int, Class)} {@link Reader#readExcelData(int, Object[], Class)}
 * 	3.invoke refresh to reset read position of current excel file.(from head of sheet) {@link Reader#refresh()}
 * 	4.invoke close to close current excel file. {@link Reader#close()}
 * </pre>
 * <p>
 * Created by xuyh at 2018/2/12 16:07.
 */
public interface Reader extends Closeable {
	/**
	 * Get data count at sheet sheetNumber
	 *
	 * @param sheetNumber sheet number
	 * @return data count
	 */
	int readExcelDataCount(int sheetNumber);

	/**
	 * Read type data at sheet sheetNumber.
	 *
	 * @param sheetNumber sheet number
	 * @param readSize    given size to read
	 * @param clazz       given type
	 * @return type data list of given size
	 */
	<T> List<T> readExcelData(int sheetNumber, int readSize, Class<T> clazz);

	/**
	 * Read type data at sheet sheetNumber and filling data into given type data array.
	 *
	 * @param sheetNumber sheet number
	 * @param ts          given model array [not null]
	 * @param clazz       given type
	 * @return read size
	 */
	<T> int readExcelData(int sheetNumber, T[] ts, Class<T> clazz);

	/**
	 * Refresh reader index of the excel file sheet position (row number).
	 * <p>
	 * <pre>
	 *     Using for re-read the excel file (all sheet) from the begin row.
	 * </pre>
	 *
	 * @return true/false
	 */
	boolean refresh();
}
