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
	 * get class type by name (entire name)
	 * <p>
	 * <pre>
	 *     For basic types :
	 *      int ---> {@link java.lang.Integer}
	 *      float ---> {@link java.lang.Float}
	 *      byte ---> {@link java.lang.Byte}
	 *      double ---> {@link java.lang.Double}
	 *      boolean ---> {@link java.lang.Boolean}
	 *      char ---> {@link java.lang.Character}
	 *      long ---> {@link java.lang.Long}
	 *      short ---> {@link java.lang.Short}
	 * </pre>
	 *
	 * @param name entire name of a type
	 * @return class type
	 */
	public static Class<?> getClassByName(String name) {
		Class<?> clazz;
		switch (name) {
		case "int":
			return Integer.class;
		case "float":
			return Float.class;
		case "byte":
			return Byte.class;
		case "double":
			return Double.class;
		case "boolean":
			return Boolean.class;
		case "char":
			return Character.class;
		case "long":
			return Long.class;
		case "short":
			return Short.class;
		}
		try {
			clazz = Class.forName(name);
		} catch (Exception e) {
			clazz = null;
		}
		return clazz;
	}

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
	 * get entire name of given clazz like "java.lang.String"
	 *
	 * @param clazz given clazz
	 * @return name String
	 */
	public static String getClassNameEntire(Class<?> clazz) {
		if (clazz == null)
			return "";
		return clazz.getName();
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
		String allName = field.getType().getName();
		if (allName.isEmpty())
			return "";
		if (!allName.contains("."))
			return allName;
		return allName.substring(allName.lastIndexOf(".") + 1, allName.length());
	}

	/**
	 * get entire name of given class like "java.lang.Integer"
	 *
	 * @param field given field
	 * @return name String
	 */
	public static String getFieldTypeNameEntire(Field field) {
		if (field == null)
			return "";
		return field.getType().getName();
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
