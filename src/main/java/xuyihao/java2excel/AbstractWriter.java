package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Model;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.core.operation.Export;
import xuyihao.java2excel.util.*;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by xuyh at 2018/1/5 10:21.
 */
public abstract class AbstractWriter {
	private String language = "en_US";

	/**
	 * Get writer language
	 *
	 * @return language
	 */
	public String getLanguage() {
		return language;
	}

	/**
	 * Set writer language
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
		return Export.createExcel(workbook, sheetNumber, model, language);
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
		return Export.insertExcelData(workbook, sheetNumber, startRowNumber, models);
	}

	/**
	 * Write existing workbook into a given file. Create a new file ignoring whether file exists.
	 *
	 * @param workbook     given workbook
	 * @param filePathName given file path name
	 * @return true/false
	 * @throws Exception exceptions
	 */
	protected boolean flush(Workbook workbook, String filePathName) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook is null!");
		File file = FileUtils.createFileDE(filePathName);
		return Common.writeFileToDisk(workbook, file);
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
