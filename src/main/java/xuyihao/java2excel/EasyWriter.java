package xuyihao.java2excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * EasyWriter
 *
 * <pre>
 *  Quickly write information into excel file at sheet 0.
 * </pre>
 * Created by Xuyh at 2017/12/30 下午 03:15.
 */
public class EasyWriter extends AbstractWriter {
	private static Logger logger = LoggerFactory.getLogger(EasyWriter.class);

	/**
	 * Write excel template file using type clazz.
	 *
	 * @param clazz             model type
	 * @param excelFilepathName excel file path name
	 * @return
	 */
	public boolean writeExcelTemplate(Class<?> clazz, String excelFilepathName) {
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook();
			if (!writeTemplate(clazz, workbook, 0))
				return false;
			if (!flush(workbook, excelFilepathName))
				return false;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		} finally {
			close(workbook);
		}
		return true;
	}

	/**
	 * Write excel file (including template and data).
	 *
	 * @param tList             data list
	 * @param excelFilepathName excel file path name
	 * @return
	 */
	public boolean writeExcelDataAll(List<?> tList, String excelFilepathName) {
		if (tList == null || tList.isEmpty())
			return false;
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook();
			if (!writeTemplate(tList.get(0).getClass(), workbook, 0))
				return false;
			if (!writeData(tList, workbook, 0, 0))
				return false;
			if (!flush(workbook, excelFilepathName))
				return false;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		} finally {
			close(workbook);
		}
		return true;
	}
}
