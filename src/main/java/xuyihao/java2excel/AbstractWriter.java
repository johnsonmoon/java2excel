package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Template;
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
	private String templateLanguage = "en_US";

	/**
	 * Get template language
	 *
	 * @return language
	 */
	public String getTemplateLanguage() {
		return templateLanguage;
	}

	/**
	 * Set template language
	 *
	 * @param templateLanguage language (en_US, zh_CN .etc)
	 */
	public void setTemplateLanguage(String templateLanguage) {
		this.templateLanguage = templateLanguage;
	}

	/**
	 * Generate excel data template instance from given class.
	 * <p>
	 * <pre>
	 *     If given class has not declared @Template or @Attribute annotations,
	 *     generate from reflection only. Otherwise generate by annotation declaration
	 *     and reflection.
	 * </pre>
	 *
	 * @param tClass given class
	 * @see xuyihao.java2excel.core.entity.annotation.Template
	 * @see xuyihao.java2excel.core.entity.annotation.Attribute
	 */
	public Template generateTemplate(Class<?> tClass) {
		if (!AnnotationUtils.hasAnnotationTemplate(tClass)) {
			return generateTemplateWithReflectionOnly(tClass);
		} else {
			return generateTemplateWithAnnotation(tClass);
		}
	}

	/**
	 * Generate template with name and fields of the given clazz.
	 *
	 * @param clazz given clazz
	 */
	private Template generateTemplateWithReflectionOnly(Class<?> clazz) {
		Template template = new Template();
		template.setName(ReflectionUtils.getClassNameShort(clazz));
		template.setJavaClassName(ReflectionUtils.getClassNameEntire(clazz));
		for (Field field : ReflectionUtils.getFieldsUnStaticUnFinal(clazz)) {
			Attribute attribute = new Attribute();
			attribute.setAttrCode(field.getName());
			attribute.setAttrName(field.getName());
			attribute.setAttrType(ReflectionUtils.getFieldTypeNameShort(field));
			attribute.setJavaClassName(ReflectionUtils.getFieldTypeNameEntire(field));
			template.addAttribute(attribute);
		}
		return template;
	}

	/**
	 * Generate template with declared annotations and reflection.
	 *
	 * @param clazz given clazz.
	 * @see xuyihao.java2excel.core.entity.annotation.Template
	 * @see xuyihao.java2excel.core.entity.annotation.Attribute
	 */
	private Template generateTemplateWithAnnotation(Class<?> clazz) {
		if (!AnnotationUtils.hasAnnotationTemplate(clazz))
			return null;
		Template template = new Template();
		xuyihao.java2excel.core.entity.annotation.Template templateAnnotation = AnnotationUtils
				.getAnnotationTemplate(clazz);
		if (templateAnnotation == null)
			return null;
		template.setName(templateAnnotation.name());
		template.setJavaClassName(ReflectionUtils.getClassNameEntire(clazz));
		for (Field field : ReflectionUtils.getFieldsAll(clazz)) {
			if (!AnnotationUtils.hasAnnotationAttribute(field))
				continue;
			xuyihao.java2excel.core.entity.annotation.Attribute attributeAnnotation = AnnotationUtils
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
			template.addAttribute(attribute);
		}
		return template;
	}

	/**
	 * Generate template data from given type instance list.
	 *
	 * <pre>
	 * 	Attribute value. Complex data(map, list .etc) with json format.(For read data from excel into objects).
	 * </pre>
	 *
	 * @param tList given type instance list.
	 * @return template data list
	 */
	private List<Template> generateData(List<?> tList) {
		List<Template> templates = new ArrayList<>();
		if (tList == null || tList.isEmpty())
			return templates;
		Class<?> clazz = tList.get(0).getClass();
		Template template = generateTemplate(clazz);
		if (template == null)
			return templates;
		for (Object t : tList) {
			Template data = new Template();
			data.setName(template.getName());
			data.setJavaClassName(template.getJavaClassName());
			data.setAttributes(template.getAttributes());
			for (Attribute attribute : data.getAttributes()) {
				Object value = ReflectionUtils.getFieldValue(t, attribute.getAttrCode());
				data.addAttrValue(ValueUtils.formatValue(value));
			}
			templates.add(data);
		}
		return templates;
	}

	/**
	 * Write template into a workbook's given sheet.
	 *
	 * @param clazz       given clazz to generate template.
	 * @param workbook    given workbook.
	 * @param sheetNumber given sheet number.
	 * @return true/false
	 */
	protected boolean writeTemplate(Class<?> clazz, Workbook workbook, int sheetNumber) {
		Template template = generateTemplate(clazz);
		return Export.createExcel(workbook, sheetNumber, template, templateLanguage);
	}

	/**
	 * Write template data into a workbook's given sheet starting from given start row number.
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
		List<Template> templates = generateData(tList);
		return Export.insertExcelData(workbook, sheetNumber, startRowNumber, templates);
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
