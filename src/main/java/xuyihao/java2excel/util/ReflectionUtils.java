package xuyihao.java2excel.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by xuyh at 2018/1/3 10:46.
 */
public class ReflectionUtils {
	/**
	 * set field value for the given object with given field name
	 *
	 * @param o     given object
	 * @param name  given field name
	 * @param value field value to be set
	 * @return true/false succeeded or failed
	 */
	public static boolean setFieldValue(Object o, String name, Object value) {
		boolean result;
		try {
			Field field = o.getClass().getDeclaredField(name);
			makeFieldAccessible(field);
			field.set(o, value);
			result = true;
		} catch (Exception e) {
			result = false;
		}
		return result;
	}

	/**
	 * get field value of given object with given field name
	 *
	 * @param o    given object
	 * @param name given field name
	 * @return field value (null if field not exists)
	 */
	public static Object getFieldValue(Object o, String name) {
		Object value;
		try {
			Field field = o.getClass().getDeclaredField(name);
			makeFieldAccessible(field);
			value = field.get(o);
		} catch (Exception e) {
			value = null;
		}
		return value;
	}

	/**
	 * get short name of given clazz like "Student"
	 *
	 * @param clazz given clazz
	 * @return name String
	 */
	public static String getClassNameShort(Class<?> clazz) {
		if (clazz == null)
			return "";
		String allName = clazz.getName();
		if (allName == null || allName.isEmpty())
			return "";
		if (!allName.contains("."))
			return allName;
		return allName.substring(allName.lastIndexOf(".") + 1, allName.length());
	}

	/**
	 * get short name of given class like "Integer"
	 *
	 * @param field given field
	 * @return name String
	 */
	public static String getFieldTypeNameShort(Field field) {
		if (field == null)
			return "";
		String allName = field.getType().toString();
		if (allName.isEmpty())
			return "";
		if (!allName.contains("."))
			return allName;
		return allName.substring(allName.lastIndexOf(".") + 1, allName.length());
	}

	/**
	 * get declared fields of given clazz which is unStatic and unFinal
	 *
	 * @param clazz given clazz
	 * @return list of fields
	 */
	public static List<Field> getFieldsUnStaticUnFinal(Class<?> clazz) {
		List<Field> fieldList = new ArrayList<>();
		if (clazz == null)
			return fieldList;
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			if (isFieldStatic(field) || isFieldFinal(field))
				continue;
			fieldList.add(field);
		}
		return fieldList;
	}

	/**
	 * get declared fields of given clazz
	 *
	 * @param clazz given clazz
	 * @return list of fields
	 */
	public static List<Field> getFieldsAll(Class<?> clazz) {
		List<Field> fieldList = new ArrayList<>();
		if (clazz == null)
			return fieldList;
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			fieldList.add(field);
		}
		return fieldList;
	}

	/**
	 * Determine whether the given field is a "public static final *" constant.
	 *
	 * @param field the field to check
	 */
	public static boolean isFieldPublicStaticFinal(Field field) {
		int modifiers = field.getModifiers();
		return (Modifier.isPublic(modifiers) && Modifier.isStatic(modifiers) && Modifier.isFinal(modifiers));
	}

	/**
	 * Determine whether the given field is a "* static final *" constant.
	 *
	 * @param field the field to check
	 */
	public static boolean isFieldStaticFinal(Field field) {
		int modifiers = field.getModifiers();
		return (Modifier.isStatic(modifiers) && Modifier.isFinal(modifiers));
	}

	/**
	 * Determine whether the given field is a "* static *" constant.
	 *
	 * @param field the field to check
	 */
	public static boolean isFieldStatic(Field field) {
		int modifiers = field.getModifiers();
		return Modifier.isStatic(modifiers);
	}

	/**
	 * Determine whether the given field is a "* final *" constant.
	 *
	 * @param field the field to check
	 */
	public static boolean isFieldFinal(Field field) {
		int modifiers = field.getModifiers();
		return Modifier.isFinal(modifiers);
	}

	/**
	 * Determine whether the given field is a "public *" constant.
	 *
	 * @param field the field to check
	 */
	public static boolean isFieldPublic(Field field) {
		int modifiers = field.getModifiers();
		return Modifier.isPublic(modifiers);
	}

	/**
	 * Determine whether the given field is a "private *" constant.
	 *
	 * @param field the field to check
	 */
	public static boolean isFieldPrivate(Field field) {
		int modifiers = field.getModifiers();
		return Modifier.isPrivate(modifiers);
	}

	/**
	 * Determine whether the given field is a "protected *" constant.
	 *
	 * @param field the field to check
	 */
	public static boolean isFieldProtected(Field field) {
		int modifiers = field.getModifiers();
		return Modifier.isProtected(modifiers);
	}

	/**
	 * Make the given field accessible, explicitly setting it accessible if
	 * necessary. The {@code setAccessible(true)} method is only called
	 * when actually necessary, to avoid unnecessary conflicts with a JVM
	 * SecurityManager (if active).
	 *
	 * @param field the field to make accessible
	 * @see java.lang.reflect.Field#setAccessible
	 */
	public static void makeFieldAccessible(Field field) {
		if ((!Modifier.isPublic(field.getModifiers()) ||
				!Modifier.isPublic(field.getDeclaringClass().getModifiers()) ||
				Modifier.isFinal(field.getModifiers())) && !field.isAccessible()) {
			field.setAccessible(true);
		}
	}

	/**
	 * Make the given constructor accessible, explicitly setting it accessible
	 * if necessary. The {@code setAccessible(true)} method is only called
	 * when actually necessary, to avoid unnecessary conflicts with a JVM
	 * SecurityManager (if active).
	 *
	 * @param ctor the constructor to make accessible
	 * @see java.lang.reflect.Constructor#setAccessible
	 */
	public static void makeConstructorAccessible(Constructor<?> ctor) {
		if ((!Modifier.isPublic(ctor.getModifiers()) ||
				!Modifier.isPublic(ctor.getDeclaringClass().getModifiers())) && !ctor.isAccessible()) {
			ctor.setAccessible(true);
		}
	}

	/**
	 * Make the given method accessible, explicitly setting it accessible if
	 * necessary. The {@code setAccessible(true)} method is only called
	 * when actually necessary, to avoid unnecessary conflicts with a JVM
	 * SecurityManager (if active).
	 *
	 * @param method the method to make accessible
	 * @see java.lang.reflect.Method#setAccessible
	 */
	public static void makeMethodAccessible(Method method) {
		if ((!Modifier.isPublic(method.getModifiers()) ||
				!Modifier.isPublic(method.getDeclaringClass().getModifiers())) && !method.isAccessible()) {
			method.setAccessible(true);
		}
	}
}
