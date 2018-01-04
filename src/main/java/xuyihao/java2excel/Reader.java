package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Template;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.core.operation.Import;
import xuyihao.java2excel.util.JsonUtils;
import xuyihao.java2excel.util.ReflectionUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2017/12/30 下午 03:15.
 */
@SuppressWarnings("all")
public class Reader {
	private static Logger logger = LoggerFactory.getLogger(Reader.class);
	private Template template;

	/**
	 * read type info from a given excel file path (if type info exists)
	 *
	 * @param excelFilepathName excel file path name
	 * @return type (nullable)
	 */
	public Class<?> readExcelJavaClass(String excelFilepathName) {
		Workbook workbook = null;
		try {
			File file = new File(excelFilepathName);
			if (!file.exists())
				return null;
			workbook = new XSSFWorkbook(file);
			template = Import.getTemplateFromExcel(workbook, 0);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return null;
		} finally {
			Common.closeWorkbook(workbook);
		}
		if (template == null)
			return null;
		String type = template.getJavaClassName();
		if (type == null)
			return null;
		return ReflectionUtils.getClassByName(type);
	}

	/**
	 * read template return from a given excel file path
	 *
	 * @param excelFilepathName excel file path name
	 * @return template (nullable)
	 */
	public Template readExcelTemplate(String excelFilepathName) {
		Workbook workbook = null;
		try {
			File file = new File(excelFilepathName);
			if (!file.exists())
				return null;
			workbook = new XSSFWorkbook(file);
			template = Import.getTemplateFromExcel(workbook, 0);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return null;
		} finally {
			Common.closeWorkbook(workbook);
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
		return readAll(excelFilePathName);
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
		List<Template> templateList = readAll(excelFilePathName);
		if (template == null)
			return dataList;
		if (template.getJavaClassName() == null || template.getJavaClassName().isEmpty())
			return dataList;
		if (!template.getJavaClassName().equals(ReflectionUtils.getClassNameEntire(clazz)))
			return dataList;
		if (templateList == null || templateList.isEmpty()) {
			return dataList;
		}
		for (Template temp : templateList) {
			Object object = ReflectionUtils.newObjectInstance(clazz);
			if (object == null)
				continue;
			List<Attribute> attributes = temp.getAttributes();
			List<String> attrValues = temp.getAttrValues();
			if (attributes == null || attributes.isEmpty())
				continue;
			if (attrValues == null || attrValues.isEmpty())
				continue;
			if (attributes.size() != attrValues.size())
				continue;
			for (int i = 0; i < attributes.size(); i++) {
				Attribute attribute = attributes.get(i);
				String fieldName = attribute.getAttrCode();
				if (fieldName == null)
					continue;
				String fieldJavaType = attribute.getJavaClassName();
				if (fieldJavaType == null)
					continue;
				String fieldStrValue = attrValues.get(i);
				if (fieldStrValue == null)
					continue;
				Object fieldRealValue = formatValue(fieldStrValue, ReflectionUtils.getClassByName(fieldJavaType));
				if (fieldRealValue == null)
					continue;
				ReflectionUtils.setFieldValue(object, fieldName, fieldRealValue);
			}
			try {
				dataList.add((T) object);
			} catch (Exception e) {
				logger.warn(e.getMessage(), e);
			}
		}
		return dataList;
	}

	private List<Template> readAll(String excelFilePathName) {
		List<Template> templates = new ArrayList<>();
		Workbook workbook = null;
		try {
			File file = new File(excelFilePathName);
			if (!file.exists())
				return templates;
			workbook = new XSSFWorkbook(file);
			template = Import.getTemplateFromExcel(workbook, 0);
			if (template == null)
				return templates;
			int count = Import.getDataCountFromExcel(workbook, 0);
			if (count == 0)
				return templates;
			int pageSize = 10;
			for (int page = 0; page <= count / pageSize; page++) {
				templates.addAll(Import.getDataFromExcel(workbook, 0, page * pageSize, pageSize));
			}
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return templates;
		} finally {
			Common.closeWorkbook(workbook);
		}
		return templates;
	}

	/**
	 * Format string instance into given type object.
	 * <p>
	 * <pre>
	 *     from String format   <--- Integer, Float, Byte, Double, Boolean, Character, Long, Short, String
	 *     from json format   <--- Other type
	 * </pre>
	 *
	 * @param value String instance to be parsed
	 * @param clazz object type String instance to be parsed
	 * @param <T>   type
	 * @return object parsed
	 */
	public static <T> Object formatValue(String value, Class<T> clazz) {
		if (value == null)
			return null;

		if (clazz == null)
			return null;

		if (clazz == Float.class) {
			return Float.parseFloat(value);
		}

		if (clazz == Byte.class) {
			return Byte.parseByte(value);
		}

		if (clazz == Double.class) {
			return Double.parseDouble(value);
		}

		if (clazz == Boolean.class) {
			return Boolean.parseBoolean(value);
		}

		if (clazz == Character.class) {
			char[] chars = value.toCharArray();
			if (chars.length == 0) {
				return null;
			}
			return chars[0];
		}

		if (clazz == Long.class) {
			return Long.parseLong(value);
		}

		if (clazz == Short.class) {
			return Short.parseShort(value);
		}

		if (clazz == String.class) {
			return value;
		}

		return JsonUtils.JsonStr2Obj(value, clazz);
	}
}
