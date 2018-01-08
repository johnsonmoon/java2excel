package xuyihao.java2excel.core.entity.dict;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by xuyh at 2017年12月29日 上午11:04:00.
 */
public enum Meta {
	FILEDS("en_US:fields", "zh_CN:字段"), //fields
	DONT_MODIFY("en_US:do not modify this row", "zh_CN:请勿编辑此行"), //do not modify this row
	DATA_FORMAT("en_US:data format", "zh_CN:数据格式"), //data format
	DEFAULT_VALUE("en_US:default value", "zh_CN:默认值"), //default value
	DATA("en_US:data", "zh_CN:数据"), //data
	DATA_FLAG("en_US:data flag", "zh_CN:数据标识");//data flag

	private Map<String, String> languageNameMap = new HashMap<>();

	Meta(String... languageNameStrs) {
		for (String languageNameStr : languageNameStrs) {
			String[] array = languageNameStr.split(":");
			languageNameMap.put(array[0], array[1]);
		}
	}

	/**
	 * get meta info of language
	 *
	 * @param language language
	 * @return meta info
	 */
	public String getMeta(String language) {
		if (supportLanguage(language))
			return languageNameMap.get(language);
		return languageNameMap.get("en_US");
	}

	/**
	 * is meta support the language
	 *
	 * @param language language
	 * @return true/false
	 */
	public boolean supportLanguage(String language) {
		return languageNameMap.containsKey(language);
	}
}
