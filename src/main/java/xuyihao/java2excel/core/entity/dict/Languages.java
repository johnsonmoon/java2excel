package xuyihao.java2excel.core.entity.dict;

/**
 * Created by xuyh at 2018/1/10 17:34.
 */
public enum Languages {
	EN_US("en_US"), ZH_CN("zh_CN");

	private String language;

	Languages(String language) {
		this.language = language;
	}

	public String getLanguage() {
		return language;
	}
}
