package xuyihao.java2excel.util;

import xuyihao.java2excel.core.entity.custom.annotation.Column;
import xuyihao.java2excel.core.entity.formatted.model.annotation.Attribute;
import xuyihao.java2excel.core.entity.formatted.model.annotation.Model;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;

/**
 * Created by xuyh at 2018/1/3 11:43.
 */
public class AnnotationUtils {
	/**
	 * get declared annotations of the given class
	 *
	 * @param clazz given class
	 * @return annotations
	 */
	public static Annotation[] getClassAnnotations(Class<?> clazz) {
		return clazz.getAnnotations();
	}

	/**
	 * Whether the given class has annotation @Model
	 *
	 * @param clazz given class
	 * @return true/false
	 */
	public static boolean hasAnnotationModel(Class<?> clazz) {
		Annotation[] annotations = clazz.getAnnotations();
		if (annotations == null || annotations.length == 0)
			return false;
		for (Annotation annotation : annotations) {
			if (annotation instanceof Model)
				return true;
		}
		return false;
	}

	/**
	 * Get annotation @Model instance of the given class if exists.
	 *
	 * @param clazz given class
	 * @return @Model annotation instance
	 */
	public static Model getAnnotationModel(Class<?> clazz) {
		if (clazz == null)
			return null;
		Annotation[] annotations = clazz.getAnnotations();
		if (annotations == null || annotations.length == 0)
			return null;
		for (Annotation annotation : annotations) {
			if (annotation instanceof Model)
				return (Model) annotation;
		}
		return null;
	}

	/**
	 * get declared annotations of the given field
	 *
	 * @param field given field
	 * @return annotations
	 */
	public static Annotation[] getFieldAnnotations(Field field) {
		return field.getAnnotations();
	}

	/**
	 * Whether the given field has annotation @Attribute
	 *
	 * @param field given field
	 * @return true/false
	 */
	public static boolean hasAnnotationAttribute(Field field) {
		Annotation[] annotations = field.getAnnotations();
		if (annotations == null || annotations.length == 0)
			return false;
		for (Annotation annotation : annotations) {
			if (annotation instanceof Attribute)
				return true;
		}
		return false;
	}

	/**
	 * Get annotation @Attribute instance of the given field if exists.
	 *
	 * @param field given field
	 * @return @Attribute annotation instance
	 */
	public static Attribute getAnnotationAttribute(Field field) {
		if (field == null)
			return null;
		Annotation[] annotations = field.getAnnotations();
		if (annotations == null || annotations.length == 0)
			return null;
		for (Annotation annotation : annotations) {
			if (annotation instanceof Attribute)
				return (Attribute) annotation;
		}
		return null;
	}

	/**
	 * Whether the given field has annotation @Column
	 *
	 * @param field given field
	 * @return true/false
	 */
	public static boolean hasAnnotationColumn(Field field) {
		Annotation[] annotations = field.getAnnotations();
		if (annotations == null || annotations.length == 0)
			return false;
		for (Annotation annotation : annotations) {
			if (annotation instanceof Column)
				return true;
		}
		return false;
	}

	/**
	 * Whether the given fields has annotation @Column
	 *
	 * @param fields given fields
	 * @return true/false
	 */
	public static boolean hasAnnotationColumn(Field[] fields) {
		boolean flag = false;
		if (fields == null || fields.length == 0)
			return false;
		for (Field field : fields) {
			Annotation[] annotations = field.getAnnotations();
			if (annotations == null || annotations.length == 0)
				continue;
			for (Annotation annotation : annotations) {
				if (annotation instanceof Column)
					flag = true;
			}
		}
		return flag;
	}

	/**
	 * Get annotation @Column instance of the given field if exists.
	 *
	 * @param field given field
	 * @return @Column annotation instance
	 */
	public static Column getAnnotationColumn(Field field) {
		if (field == null)
			return null;
		Annotation[] annotations = field.getAnnotations();
		if (annotations == null || annotations.length == 0) {
			return null;
		}
		for (Annotation annotation : annotations) {
			if (annotation instanceof Column)
				return (Column) annotation;
		}
		return null;
	}
}
