package xuyihao.java2excel.util;

import java.util.Date;

/**
 * Utilities for excel cell value
 * <p>
 * Created by xuyh at 2018/1/5 10:30.
 */
public class ValueUtils {
	/**
	 * Parse string instance into given type object.
	 * <p>
	 * <pre>
	 *     from String format   <--- Integer, Float, Byte, Double, Boolean, Character, Long, Short, String, Date
	 *     from json format   <--- Other type
	 * </pre>
	 *
	 * @param value String instance to be parsed
	 * @param clazz object type String instance to be parsed
	 * @param <T>   type
	 * @return object parsed
	 */
	public static <T> Object parseValue(String value, Class<T> clazz) {
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

		if (clazz == Date.class) {
			return DateUtils.parseDateTime(value);
		}

		return JsonUtils.JsonStr2Obj(value, clazz);
	}

	/**
	 * Format object to String instance.
	 * <p>
	 * <pre>
	 *     Integer, Float, Byte, Double, Boolean, Character, Long, Short, String, Date ---> String instance
	 *     Other type ---> json format String instance
	 * </pre>
	 *
	 * @param object object to be format
	 * @return formatted String instance
	 */
	public static String formatValue(Object object) {
		if (object == null)
			return "";
		if (object instanceof Integer
				|| object instanceof Float
				|| object instanceof Byte
				|| object instanceof Double
				|| object instanceof Boolean
				|| object instanceof Character
				|| object instanceof Long
				|| object instanceof Short
				|| object instanceof String)
			return String.valueOf(object);
		if (object instanceof Date)
			return DateUtils.formatDateTime((Date) object);
		return JsonUtils.obj2JsonStr(object);
	}
}
