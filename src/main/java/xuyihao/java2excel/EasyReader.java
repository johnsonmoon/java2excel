package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.model.Template;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * EasyReader
 *
 * <pre>
 * 	Quickly read excel file at sheet 0.
 * </pre>
 * Created by Xuyh at 2017/12/30 下午 03:15.
 */
public class EasyReader extends AbstractReader {
	private static Logger logger = LoggerFactory.getLogger(EasyReader.class);

	/**
	 * read type info from a given excel file path (if type info exists)
	 *
	 * @param excelFilepathName excel file path name
	 * @return type (nullable)
	 */
	public Class<?> readExcelJavaClass(String excelFilepathName) {
		Workbook workbook = null;
		Class<?> clazz = null;
		try {
			File file = new File(excelFilepathName);
			if (!file.exists())
				return null;
			workbook = new XSSFWorkbook(file);
			clazz = readJavaClass(workbook, 0);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		} finally {
			close(workbook);
		}
		return clazz;
	}

	/**
	 * read template return from a given excel file path
	 *
	 * @param excelFilepathName excel file path name
	 * @return template (nullable)
	 */
	public Template readExcelTemplate(String excelFilepathName) {
		Workbook workbook = null;
		Template template = null;
		try {
			File file = new File(excelFilepathName);
			if (!file.exists())
				return null;
			workbook = new XSSFWorkbook(file);
			template = readTemplate(workbook, 0);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		} finally {
			close(workbook);
		}
		return template;
	}

	/**
	 * read template data from a given excel file path
	 *
	 * @param excelFilePathName excel file path name
	 * @return template data list
	 */
	public List<Template> readExcelDataAll(String excelFilePathName) {
		List<Template> templates = new ArrayList<>();
		Workbook workbook = null;
		try {
			File file = new File(excelFilePathName);
			if (!file.exists())
				return templates;
			workbook = new XSSFWorkbook(file);
			int count = readDataCount(workbook, 0);
			if (count == 0)
				return templates;
			int pageSize = 10;
			for (int page = 0; page <= count / pageSize; page++) {
				templates.addAll(readData(workbook, 0, page * pageSize, pageSize));
			}
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		} finally {
			close(workbook);
		}
		return templates;
	}

	/**
	 * read type data from a given excel file path
	 *
	 * @param excelFilePathName excel file path name
	 * @param clazz             given type
	 * @param <T>               given type
	 * @return type data list
	 */
	public <T> List<T> readExcelDataAll(String excelFilePathName, Class<T> clazz) {
		List<T> dataList = new ArrayList<>();
		Workbook workbook = null;
		try {
			File file = new File(excelFilePathName);
			if (!file.exists())
				return dataList;
			workbook = new XSSFWorkbook(file);
			int count = readDataCount(workbook, 0);
			if (count == 0)
				return dataList;
			int pageSize = 10;
			for (int page = 0; page <= count / pageSize; page++) {
				dataList.addAll(readData(clazz, workbook, 0, page * pageSize, pageSize));
			}
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		} finally {
			close(workbook);
		}
		return dataList;
	}
}
