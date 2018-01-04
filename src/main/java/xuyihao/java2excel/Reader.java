package xuyihao.java2excel;

import xuyihao.java2excel.util.JsonUtils;

/**
 * Created by Xuyh at 2017/12/30 下午 03:15.
 */
public class Reader {

	/**
	 * Format string instance into given type object.
	 * <p>
	 * <pre>
	 *     from String format   <--- Integer, Float, Byte, Double, Boolean, Character, Long, Short, String
	 *     from json format   <--- Other type
	 * </pre>
	 *
	 * @param value String instance to be parsed
	 * @param clazz object type String instance to be parsed
	 * @param <T>   type
	 * @return object parsed
	 */
	public static <T> Object formatValue(String value, Class<T> clazz) {
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

		return JsonUtils.JsonStr2Obj(value, clazz);
	}
}
