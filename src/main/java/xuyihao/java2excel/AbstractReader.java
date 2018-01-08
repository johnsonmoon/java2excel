package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Model;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.core.operation.Import;
import xuyihao.java2excel.util.ReflectionUtils;
import xuyihao.java2excel.util.ValueUtils;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by xuyh at 2018/1/5 10:21.
 */
@SuppressWarnings("all")
public abstract class AbstractReader {
	/**
	 * Read type class from a given workbook in sheet from a sheetNumber.
	 *
	 * @param workbook    given workbook
	 * @param sheetNumber given sheet number
	 * @return type class info
	 * @throws Exception exceptions
	 */
	protected Class<?> readJavaClass(Workbook workbook, int sheetNumber) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook can not be null.");
		if (sheetNumber < 0)
			throw new RuntimeException("sheet number can not below 0.");
		Model model = readTemplate(workbook, sheetNumber);
		if (model == null)
			return null;
		String className = model.getJavaClassName();
		if (className == null)
			return null;
		return ReflectionUtils.getClassByName(className);
	}

	/**
	 * Read template from a given workbook in sheet from a sheetNumber.
	 *
	 * @param workbook    given workbook
	 * @param sheetNumber given sheet number
	 * @return Model instance
	 * @throws Exception exceptions
	 */
	protected Model readTemplate(Workbook workbook, int sheetNumber) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook can not be null.");
		if (sheetNumber < 0)
			throw new RuntimeException("sheet number can not below 0.");
		Model model = Import.getTemplateFromExcel(workbook, sheetNumber);
		return model;
	}

	/**
	 * Read template data from a given workbook in sheet from a sheetNumber.
	 *
	 * @param workbook       given workbook
	 * @param sheetNumber    given sheet number
	 * @param beginRowNumber row number to begin reading
	 * @param readSize       row count to read
	 * @return Model list
	 * @throws Exception exceptions
	 */
	protected List<Model> readData(Workbook workbook, int sheetNumber, int beginRowNumber, int readSize)
			throws Exception {
		return readTemplates(workbook, sheetNumber, beginRowNumber, readSize);
	}

	/**
	 * Read type data from a given workbook in sheet from a sheetNumber.
	 *
	 * @param workbook       given workbook
	 * @param sheetNumber    given sheet number
	 * @param beginRowNumber row number to begin reading
	 * @param readSize       row count to read
	 * @param <T>            type
	 * @return type list
	 * @throws Exception exceptions
	 */
	protected <T> List<T> readData(Class<T> clazz, Workbook workbook, int sheetNumber, int beginRowNumber, int readSize)
			throws Exception {
		List<Model> modelList = readTemplates(workbook, sheetNumber, beginRowNumber, readSize);
		List<T> dataList = new ArrayList<>();
		Model model = readTemplate(workbook, sheetNumber);
		if (model == null)
			return dataList;
		if (model.getJavaClassName() == null || model.getJavaClassName().isEmpty())
			return dataList;
		if (!model.getJavaClassName().equals(ReflectionUtils.getClassNameEntire(clazz)))
			return dataList;
		if (modelList == null || modelList.isEmpty()) {
			return dataList;
		}
		for (Model temp : modelList) {
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
				Object fieldRealValue = ValueUtils.parseValue(fieldStrValue, ReflectionUtils.getClassByName(fieldJavaType));
				if (fieldRealValue == null)
					continue;
				ReflectionUtils.setFieldValue(object, fieldName, fieldRealValue);
			}
			try {
				dataList.add((T) object);
			} catch (Exception e) {
			}
		}
		return dataList;
	}

	private List<Model> readTemplates(Workbook workbook, int sheetNumber, int beginRowNumber, int readSize)
			throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook can not be null.");
		if (sheetNumber < 0)
			throw new RuntimeException("sheet number can not below 0.");
		if (beginRowNumber < 0)
			throw new RuntimeException("beginRowNumber can not below 0.");
		if (readSize < 0)
			throw new RuntimeException("readSize can not below 0.");
		Model model = readTemplate(workbook, sheetNumber);
		if (model == null)
			readTemplate(workbook, sheetNumber);
		List<Model> models = new ArrayList<>();
		models.addAll(Import.getDataFromExcel(workbook, sheetNumber, beginRowNumber, readSize));
		return models;
	}

	/**
	 * Read data count from a given workbook in sheet from a sheetNumber.
	 *
	 * @param workbook    given workbook
	 * @param sheetNumber given sheet number
	 * @return data count
	 * @throws Exception exceptions
	 */
	protected int readDataCount(Workbook workbook, int sheetNumber) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook can not be null.");
		if (sheetNumber < 0)
			throw new RuntimeException("sheet number can not below 0.");
		return Import.getDataCountFromExcel(workbook, sheetNumber);
	}

	/**
	 * Close workbook, release resources.
	 *
	 * @param workbook given workbook
	 * @return true/false
	 */
	protected boolean close(Workbook workbook) {
		return Common.closeWorkbook(workbook);
	}
}
