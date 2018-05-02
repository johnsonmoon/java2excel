package com.github.johnsonmoon.java2excel;

import org.junit.Test;

/**
 * Created by xuyh at 2018/2/12 17:17.
 */
public class ReaderTest {
	@Test
	public void test() {
		String filePathNameCustom = System.getProperty("user.dir") + "/src/test/resources/readerCustom.xlsx";
		Reader readerCustom = ExcelFactory.getReader(ExcelFactory.TYPE_CODE_CUSTOM, filePathNameCustom, 1);
		readerCustom.readExcelData(0, 2, User.class).forEach(System.out::println);
		System.out.println(readerCustom.readExcelDataCount(0));
		readerCustom.close();

		String filePathNameFormatted = System.getProperty("user.dir") + "/src/test/resources/readerFormatted.xlsx";
		Reader readerFormatted = ExcelFactory.getReader(ExcelFactory.TYPE_CODE_FORMATTED, filePathNameFormatted);
		System.out.println(readerFormatted.readExcelDataCount(0));
		readerCustom.readExcelData(0, 2, User.class).forEach(System.out::println);
		readerFormatted.close();
	}
}
