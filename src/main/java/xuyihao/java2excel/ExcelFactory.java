package xuyihao.java2excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.*;

/**
 * Created by xuyh at 2018/2/12 16:23.
 */
public class ExcelFactory {
	private static Logger logger = LoggerFactory.getLogger(ExcelFactory.class);

	/**
	 * Custom excel tool type.
	 */
	public static final int TYPE_CODE_CUSTOM = 0x01;
	/**
	 * Formatted excel tool type.
	 */
	public static final int TYPE_CODE_FORMATTED = 0x02;

	/**
	 * Get excel editor.
	 *
	 * @param type         tool type {@link ExcelFactory#TYPE_CODE_CUSTOM}, {@link ExcelFactory#TYPE_CODE_FORMATTED}
	 * @param filePathName excel file path name
	 * @param args         other parameters for initializing Editor instance.
	 * <pre>
	 * {@link FormattedEditor#FormattedEditor(String)}<br>
	 * {@link CustomEditor#CustomEditor(String)}<br>
	 * {@link CustomEditor#CustomEditor(String, int)} - the second parameter type int.
	 * </pre>
	 * @return excel editor {@link Editor}
	 */
	public static Editor getEditor(int type, String filePathName, Object... args) {
		switch (type) {
		case TYPE_CODE_CUSTOM:
			if (args == null || args.length == 0)
				return new CustomEditor(filePathName);
			Integer dataBeginRowNumber = null;
			try {
				dataBeginRowNumber = Integer.parseInt(String.valueOf(args[0]));
			} catch (Exception e) {
				logger.warn(e.getMessage(), e);
			}
			if (dataBeginRowNumber != null) {
				return new CustomEditor(filePathName, dataBeginRowNumber);
			}
			return new CustomEditor(filePathName);
		case TYPE_CODE_FORMATTED:
			return new FormattedEditor(filePathName);
		default:
			return new FormattedEditor(filePathName);
		}
	}

	/**
	 * Get excel reader.
	 *
	 * @param type         tool type {@link ExcelFactory#TYPE_CODE_CUSTOM}, {@link ExcelFactory#TYPE_CODE_FORMATTED}
	 * @param filePathName excel file path name
	 * @param args         other parameters for initializing Editor instance.
	 * <pre>
	 * {@link FormattedReader#FormattedReader(String)}<br>
	 * {@link CustomReader#CustomReader(String)}<br>
	 * {@link CustomReader#CustomReader(String, int)} - the second parameter type int.
	 * </pre>
	 * @return excel reader {@link Reader}
	 */
	public static Reader getReader(int type, String filePathName, Object... args) {
		switch (type) {
		case TYPE_CODE_CUSTOM:
			if (args == null || args.length == 0)
				return new CustomReader(filePathName);
			Integer dataBeginRowNumber = null;
			try {
				dataBeginRowNumber = Integer.parseInt(String.valueOf(args[0]));
			} catch (Exception e) {
				logger.warn(e.getMessage(), e);
			}
			if (dataBeginRowNumber != null) {
				return new CustomReader(filePathName, dataBeginRowNumber);
			}
			return new CustomReader(filePathName);
		case TYPE_CODE_FORMATTED:
			return new FormattedReader(filePathName);
		default:
			return new FormattedReader(filePathName);
		}
	}

	/**
	 * Get excel writer.
	 *
	 * @param type         tool type {@link ExcelFactory#TYPE_CODE_CUSTOM}, {@link ExcelFactory#TYPE_CODE_FORMATTED}
	 * @param filePathName excel file path name
	 * @return excel writer {@link Writer}
	 */
	public static Writer getWriter(int type, String filePathName) {
		switch (type) {
		case TYPE_CODE_CUSTOM:
			return new CustomWriter(filePathName);
		case TYPE_CODE_FORMATTED:
			return new FormattedWriter(filePathName);
		default:
			return new FormattedWriter(filePathName);
		}
	}
}
