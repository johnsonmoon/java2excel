package com.github.johnsonmoon.java2excel.core;

import com.github.johnsonmoon.java2excel.core.operation.Common;
import com.github.johnsonmoon.java2excel.util.FileUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;

/**
 * Created by xuyh at 2018/2/12 12:43.
 */
public class CustomAbstractWriter extends CustomAbstractEditor {
	/**
	 * Write existing workbook into a given file. Create a new file ignoring whether file exists.
	 *
	 * @param workbook     given workbook
	 * @param filePathName given file path name
	 * @return true/false
	 * @throws Exception exceptions
	 */
	@Override
	protected boolean flush(Workbook workbook, String filePathName) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook is null!");
		File file = FileUtils.createFileDE(filePathName);
		return Common.writeFileToDisk(workbook, file);
	}
}
