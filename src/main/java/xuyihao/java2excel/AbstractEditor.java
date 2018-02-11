package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.entity.dict.Languages;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Model;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.core.operation.format.FormatExporter;
import xuyihao.java2excel.core.operation.format.FormatImporter;
import xuyihao.java2excel.util.*;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by xuyh at 2018/1/10 17:56.
 */
@SuppressWarnings("all")
public abstract class AbstractEditor {
	private String language = Languages.EN_US.getLanguage();

	/**
	 * Get Editor language
	 *
	 * @return language
	 */
	public String getLanguage() {
		return language;
	}

	/**
	 * Set Editor language
	 *
	 * @param language language (en_US, zh_CN .etc)
	 */
	public void setLanguage(String language) {
		this.language = language;
	}

	/**
	 * Generate excel data model instance from given class.
	 * <p>
	 * <pre>
	 *     If given class has not declared @Model or @Attribute annotations,
	 *     generate from reflection only. Otherwise generate by annotation declaration
	 *     and reflection.
	 * </pre>
	 *
	 * @param tClass given class
	 * @see xuyihao.java2excel.core.entity.model.annotation.Model
	 * @see xuyihao.java2excel.core.entity.model.annotation.Attribute
	 */
	public Model generateModel(Class<?> tClass) {
		if (!AnnotationUtils.hasAnnotationModel(tClass)) {
			return generateModelWithReflectionOnly(tClass);
		} else {
			return generateModelWithAnnotation(tClass);
		}
	}

	/**
	 * Generate model with name and fields of the given clazz.
	 *
	 * @param clazz given clazz
	 */
	private Model generateModelWithReflectionOnly(Class<?> clazz) {
		Model model = new Model();
		model.setName(ReflectionUtils.getClassNameShort(clazz));
		model.setJavaClassName(ReflectionUtils.getClassNameEntire(clazz));
		for (Field field : ReflectionUtils.getFieldsUnStaticUnFinal(clazz)) {
			Attribute attribute = new Attribute();
			attribute.setAttrCode(field.getName());
			attribute.setAttrName(field.getName());
			attribute.setAttrType(ReflectionUtils.getFieldTypeNameShort(field));
			attribute.setJavaClassName(ReflectionUtils.getFieldTypeNameEntire(field));
			model.addAttribute(attribute);
		}
		return model;
	}

	/**
	 * Generate model with declared annotations and reflection.
	 *
	 * @param clazz given clazz.
	 * @see xuyihao.java2excel.core.entity.model.annotation.Model
	 * @see xuyihao.java2excel.core.entity.model.annotation.Attribute
	 */
	private Model generateModelWithAnnotation(Class<?> clazz) {
		if (!AnnotationUtils.hasAnnotationModel(clazz))
			return null;
		Model model = new Model();
		xuyihao.java2excel.core.entity.model.annotation.Model modelAnnotation = AnnotationUtils
				.getAnnotationModel(clazz);
		if (modelAnnotation == null)
			return null;
		model.setName(modelAnnotation.name());
		model.setJavaClassName(ReflectionUtils.getClassNameEntire(clazz));
		for (Field field : ReflectionUtils.getFieldsAll(clazz)) {
			if (!AnnotationUtils.hasAnnotationAttribute(field))
				continue;
			xuyihao.java2excel.core.entity.model.annotation.Attribute attributeAnnotation = AnnotationUtils
					.getAnnotationAttribute(field);
			if (attributeAnnotation == null)
				continue;
			Attribute attribute = new Attribute();
			attribute.setAttrCode(field.getName());
			attribute.setAttrName(StringUtils.replaceEmptyToNull(attributeAnnotation.attrName()));
			attribute.setAttrType(StringUtils.replaceEmptyToNull(attributeAnnotation.attrType()));
			attribute.setFormatInfo(StringUtils.replaceEmptyToNull(attributeAnnotation.formatInfo()));
			attribute.setDefaultValue(StringUtils.replaceEmptyToNull(attributeAnnotation.defaultValue()));
			attribute.setUnit(StringUtils.replaceEmptyToNull(attributeAnnotation.unit()));
			attribute.setJavaClassName(ReflectionUtils.getFieldTypeNameEntire(field));
			model.addAttribute(attribute);
		}
		return model;
	}

	/**
	 * Generate data from given type instance list.
	 * <p>
	 * <pre>
	 * 	Attribute value. Complex data(map, list .etc) with json format.(For read data from excel into objects).
	 * </pre>
	 *
	 * @param tList given type instance list.
	 * @return model data list
	 */
	private List<Model> generateData(List<?> tList) {
		List<Model> models = new ArrayList<>();
		if (tList == null || tList.isEmpty())
			return models;
		Class<?> clazz = tList.get(0).getClass();
		Model model = generateModel(clazz);
		if (model == null)
			return models;
		for (Object t : tList) {
			Model data = new Model();
			data.setName(model.getName());
			data.setJavaClassName(model.getJavaClassName());
			data.setAttributes(model.getAttributes());
			for (Attribute attribute : data.getAttributes()) {
				Object value = ReflectionUtils.getFieldValue(t, attribute.getAttrCode());
				data.addAttrValue(ValueUtils.formatValue(value));
			}
			models.add(data);
		}
		return models;
	}

	/**
	 * Write model into a workbook's given sheet.
	 *
	 * @param clazz       given clazz to generate model.
	 * @param workbook    given workbook.
	 * @param sheetNumber given sheet number.
	 * @return true/false
	 */
	protected boolean writeModel(Class<?> clazz, Workbook workbook, int sheetNumber) {
		Model model = generateModel(clazz);
		return FormatExporter.createExcel(workbook, sheetNumber, model, language);
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param tList          given clazz data list
	 * @param workbook       given workbook
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	protected boolean writeData(List<?> tList, Workbook workbook, int sheetNumber, int startRowNumber) {
		if (tList == null || tList.isEmpty())
			return false;
		List<Model> models = generateData(tList);
		return FormatExporter.insertExcelData(workbook, sheetNumber, startRowNumber, models);
	}

	/**
	 * Write existing workbook into a given file. Create a new file if the file isn't exist.
	 *
	 * @param workbook     given workbook
	 * @param filePathName given file path name
	 * @return true/false
	 * @throws Exception exceptions
	 */
	protected boolean flush(Workbook workbook, String filePathName) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook is null!");
		File file = FileUtils.createFile(filePathName);
		return Common.writeFileToDisk(workbook, file);
	}

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
		Model model = readModel(workbook, sheetNumber);
		if (model == null)
			return null;
		String className = model.getJavaClassName();
		if (className == null)
			return null;
		return ReflectionUtils.getClassByName(className);
	}

	/**
	 * Read model from a given workbook in sheet from a sheetNumber.
	 *
	 * @param workbook    given workbook
	 * @param sheetNumber given sheet number
	 * @return Model instance
	 * @throws Exception exceptions
	 */
	protected Model readModel(Workbook workbook, int sheetNumber) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook can not be null.");
		if (sheetNumber < 0)
			throw new RuntimeException("sheet number can not below 0.");
		Model model = FormatImporter.getModelFromExcel(workbook, sheetNumber);
		return model;
	}

	/**
	 * Read model data from a given workbook in sheet from a sheetNumber.
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
		return readModels(workbook, sheetNumber, beginRowNumber, readSize);
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
		List<Model> modelList = readModels(workbook, sheetNumber, beginRowNumber, readSize);
		List<T> dataList = new ArrayList<>();
		Model model = readModel(workbook, sheetNumber);
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

	private List<Model> readModels(Workbook workbook, int sheetNumber, int beginRowNumber, int readSize)
			throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook can not be null.");
		if (sheetNumber < 0)
			throw new RuntimeException("sheet number can not below 0.");
		if (beginRowNumber < 0)
			throw new RuntimeException("beginRowNumber can not below 0.");
		if (readSize < 0)
			throw new RuntimeException("readSize can not below 0.");
		Model model = readModel(workbook, sheetNumber);
		if (model == null)
			readModel(workbook, sheetNumber);
		List<Model> models = new ArrayList<>();
		models.addAll(FormatImporter.getDataFromExcel(workbook, sheetNumber, beginRowNumber, readSize));
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
		return FormatImporter.getDataCountFromExcel(workbook, sheetNumber);
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
