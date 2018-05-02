package com.github.johnsonmoon.java2excel.util;

import org.codehaus.jackson.map.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Created by xuyh at 2018/1/3 15:59.
 */
public class JsonUtils {
	private static Logger logger = LoggerFactory.getLogger(JsonUtils.class);

	/**
	 * Parsing json string into an object instance.
	 *
	 * @param jsonStr json string
	 * @param tClass  class type
	 * @return class type object instance
	 */
	public static <T> Object JsonStr2Obj(String jsonStr, Class<T> tClass) {
		T t = null;
		try {
			ObjectMapper objectMapper = new ObjectMapper();
			t = objectMapper.readValue(jsonStr, tClass);
		} catch (Exception e) {
			logger.warn(String.format("Exception when parse json string, [%s]", e.getMessage()));
		}
		return t;
	}

	/**
	 * Parsing Object into a json string.
	 *
	 * @param obj object instance
	 * @return json string
	 */
	public static <T> String obj2JsonStr(T obj) {
		ObjectMapper mapper = new ObjectMapper();
		String jsonStr = null;
		try {
			jsonStr = mapper.writeValueAsString(obj);
		} catch (Exception e) {
			logger.warn(String.format("Exception when parse object to jsonString, [%s]", e.getMessage()));
		}
		return jsonStr;
	}
}
