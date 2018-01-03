package xuyihao.java2excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Template;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.core.operation.Export;
import xuyihao.java2excel.util.AnnotationUtils;
import xuyihao.java2excel.util.FileUtils;
import xuyihao.java2excel.util.ReflectionUtils;
import xuyihao.java2excel.util.StringUtils;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2017/12/30 下午 03:15.
 */
public class Writer {
	private static Logger logger = LoggerFactory.getLogger(Writer.class);
	private Template template;
	private String templateLanguage = "en_US";

	public Template getTemplate() {
		return template;
	}

	public void setTemplate(Template template) {
		this.template = template;
	}

	public String getTemplateLanguage() {
		return templateLanguage;
	}

	public void setTemplateLanguage(String templateLanguage) {
		this.templateLanguage = templateLanguage;
	}

	/**
	 * Write excel template file using type clazz.
	 *
	 * @param clazz             model type
	 * @param excelFilepathName excel file path name
	 * @return
	 */
	public boolean writeExcelTemplate(Class<?> clazz, String excelFilepathName) {
		generateTemplate(clazz);
		XSSFWorkbook workbook = null;
		try {
			File file = FileUtils.createFileDE(excelFilepathName);
			if (file == null)
				return false;
			workbook = new XSSFWorkbook();
			if (!Export.createExcel(workbook, 0, template, this.templateLanguage))
				return false;
			if (!Common.writeFileToDisk(workbook, file))
				return false;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		} finally {
			Common.closeWorkbook(workbook);
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
	public boolean writeExcelData(List<?> tList, String excelFilepathName) {
		if (tList == null || tList.isEmpty())
			return false;
		generateTemplate(tList.get(0).getClass());
		XSSFWorkbook workbook = null;
		try {
			File file = FileUtils.createFileDE(excelFilepathName);
			if (file == null)
				return false;
			workbook = new XSSFWorkbook();
			if (!Export.createExcel(workbook, 0, template, this.templateLanguage))
				return false;
			if (!Export.insertExcelData(workbook, 0, 6, generateData(tList)))
				return false;
			if (!Common.writeFileToDisk(workbook, file))
				return false;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		} finally {
			Common.closeWorkbook(workbook);
		}
		return true;
	}

	private void generateTemplate(Class<?> tClass) {
		//If type has @Template annotation modifier, do generate template by annotations.
		//If type has no annotation modifiers, do generate template by reflection only.
		if (!AnnotationUtils.hasAnnotationTemplate(tClass)) {
			generateTemplateWithReflectionOnly(tClass);
		} else {
			generateTemplateWithAnnotation(tClass);
		}
	}

	private void generateTemplateWithReflectionOnly(Class<?> clazz) {
		this.template = new Template();
		template.setName(ReflectionUtils.getClassNameShort(clazz));
		for (Field field : ReflectionUtils.getFieldsUnStaticUnFinal(clazz)) {
			Attribute attribute = new Attribute();
			attribute.setAttrCode(field.getName());
			attribute.setAttrName(field.getName());
			attribute.setAttrType(ReflectionUtils.getFieldTypeNameShort(field));
			template.addAttribute(attribute);
		}
	}

	private void generateTemplateWithAnnotation(Class<?> clazz) {
		if (!AnnotationUtils.hasAnnotationTemplate(clazz))
			return;
		this.template = new Template();
		xuyihao.java2excel.core.entity.annotation.Template templateAnnotation = AnnotationUtils
				.getAnnotationTemplate(clazz);
		if (templateAnnotation == null)
			return;
		template.setName(templateAnnotation.name());
		for (Field field : ReflectionUtils.getFieldsAll(clazz)) {
			if (!AnnotationUtils.hasAnnotationAttribute(field))
				continue;
			xuyihao.java2excel.core.entity.annotation.Attribute attributeAnnotation = AnnotationUtils
					.getAnnotationAttribute(field);
			if (attributeAnnotation == null)
				continue;
			Attribute attribute = new Attribute();
			attribute.setAttrCode(StringUtils.replaceEmptyToNull(attributeAnnotation.attrCode()));
			attribute.setAttrName(StringUtils.replaceEmptyToNull(attributeAnnotation.attrName()));
			attribute.setAttrType(StringUtils.replaceEmptyToNull(attributeAnnotation.attrType()));
			attribute.setFormatInfo(StringUtils.replaceEmptyToNull(attributeAnnotation.formatInfo()));
			attribute.setDefaultValue(StringUtils.replaceEmptyToNull(attributeAnnotation.defaultValue()));
			attribute.setUnit(StringUtils.replaceEmptyToNull(attributeAnnotation.unit()));
			template.addAttribute(attribute);
		}
	}

	private List<Template> generateData(List<?> tList) {
		//TODO 对不同的数据类型添加不同的attrValue格式(推荐使用json格式)
		//Attribute value. Complex data(map, list .etc) with json format.(For read data from excel into objects).
		List<Template> templates = new ArrayList<>();
		for (Object t : tList) {
			Template template = new Template();
			template.setName(this.template.getName());
			template.setAttributes(this.template.getAttributes());
			for (Attribute attribute : template.getAttributes()) {
				Object value = ReflectionUtils.getFieldValue(t, attribute.getAttrCode());
				template.addAttrValue(value == null ? "" : String.valueOf(value));
			}
			templates.add(template);
		}
		return templates;
	}
}
