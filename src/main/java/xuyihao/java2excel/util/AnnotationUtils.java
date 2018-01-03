package xuyihao.java2excel.util;

import xuyihao.java2excel.core.entity.annotation.Attribute;
import xuyihao.java2excel.core.entity.annotation.Template;

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
	 * Whether the given class has annotation @Template
	 *
	 * @param clazz given class
	 * @return true/false
	 */
	public static boolean hasAnnotationTemplate(Class<?> clazz) {
		Annotation[] annotations = clazz.getAnnotations();
		if (annotations == null || annotations.length == 0)
			return false;
		for (Annotation annotation : annotations) {
			if (annotation instanceof Template)
				return true;
		}
		return false;
	}

	/**
	 * Get annotation @Template instance of the given class if exists.
	 *
	 * @param clazz given class
	 * @return @Template annotation instance
	 */
	public static Template getAnnotationTemplate(Class<?> clazz) {
		if (clazz == null)
			return null;
		Annotation[] annotations = clazz.getAnnotations();
		if (annotations == null || annotations.length == 0)
			return null;
		for (Annotation annotation : annotations) {
			if (annotation instanceof Template)
				return (Template) annotation;
		}
		return null;
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
}
