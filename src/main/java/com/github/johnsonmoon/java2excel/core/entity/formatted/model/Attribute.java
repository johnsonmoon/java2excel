package com.github.johnsonmoon.java2excel.core.entity.formatted.model;

/**
 * Model field
 * <p>
 * Created by Xuyh at 2016/07/22 上午 10:42.
 */
public class Attribute {
	/**
	 * {@link java.lang.reflect.Field#getName()}
	 */
	private String attrCode;//field code
	private String attrName;//field name
	private String attrType;//field type
	private String formatInfo;//data format info
	private String defaultValue;//field data default value
	private String unit;//field data value unit
	private String javaClassName;//java (class) type full name

	public Attribute() {
	}

	public Attribute(String attrCode, String attrName, String attrType, String formatInfo, String defaultValue,
			String unit) {
		this.attrCode = attrCode;
		this.attrName = attrName;
		this.attrType = attrType;
		this.formatInfo = formatInfo;
		this.defaultValue = defaultValue;
		this.unit = unit;
	}

	public String getFormatInfo() {
		return formatInfo;
	}

	public void setFormatInfo(String formatInfo) {
		this.formatInfo = formatInfo;
	}

	/**
	 * attrCode - fieldName
	 * <p>
	 * {@link java.lang.reflect.Field#getName()}
	 */
	public String getAttrCode() {
		return attrCode;
	}

	public void setAttrCode(String attrCode) {
		this.attrCode = attrCode;
	}

	public String getAttrName() {
		return attrName;
	}

	public void setAttrName(String attrName) {
		this.attrName = attrName;
	}

	public String getAttrType() {
		return attrType;
	}

	public void setAttrType(String attrType) {
		this.attrType = attrType;
	}

	public String getDefaultValue() {
		return defaultValue;
	}

	public void setDefaultValue(String defaultValue) {
		this.defaultValue = defaultValue;
	}

	public String getUnit() {
		return unit;
	}

	public void setUnit(String unit) {
		this.unit = unit;
	}

	public String getJavaClassName() {
		return javaClassName;
	}

	public void setJavaClassName(String javaClassName) {
		this.javaClassName = javaClassName;
	}

	@Override
	public String toString() {
		return "Attribute{" +
				"attrCode='" + attrCode + '\'' +
				", attrName='" + attrName + '\'' +
				", attrType='" + attrType + '\'' +
				", formatInfo='" + formatInfo + '\'' +
				", defaultValue='" + defaultValue + '\'' +
				", unit='" + unit + '\'' +
				", javaClassName='" + javaClassName + '\'' +
				'}';
	}
}
