package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.util.FileUtils;

import java.io.File;

/**
 * Created by xuyh at 2018/3/20 16:45.
 */
public abstract class FreeAbstractWriter extends FreeAbstractEditor {
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
