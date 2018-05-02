package com.github.johnsonmoon.java2excel.util;

/**
 * Created by xuyh at 2018/1/3 15:50.
 */
public class StringUtils {
	/**
	 * if value equals "" return null
	 *
	 * @param value given value
	 * @return null if given value equals ""
	 */
	public static String replaceEmptyToNull(String value) {
		if (value == null)
			return null;
		if (value.equals(""))
			return null;
		return value;
	}

	/**
	 * if value equals null return ""
	 *
	 * @param value given value
	 * @return "" if value equals null
	 */
	public static String replaceNullToEmpty(String value) {
		if (value == null)
			return "";
		return value;
	}
}
