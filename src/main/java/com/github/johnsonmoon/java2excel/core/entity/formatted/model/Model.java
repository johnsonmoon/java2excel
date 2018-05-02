package com.github.johnsonmoon.java2excel.core.entity.formatted.model;

import java.util.ArrayList;
import java.util.List;

/**
 * Data model
 * <p>
 * Created by Xuyh at 2016/07/21 下午 01:40.
 */
public class Model {
	private String name;//model name (Sheet name)
	private String javaClassName;//java entity (class) type full name
	private List<Attribute> attributes;//fields
	private List<String> attrValues;//fields' value

	public Model() {
	}

	public Model(String name, List<Attribute> attributes, List<String> attrValues) {
		this.name = name;
		this.attributes = attributes;
		this.attrValues = attrValues;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getJavaClassName() {
		return javaClassName;
	}

	public void setJavaClassName(String javaClassName) {
		this.javaClassName = javaClassName;
	}

	public List<Attribute> getAttributes() {
		return attributes;
	}

	public void setAttributes(List<Attribute> attributes) {
		this.attributes = attributes;
	}

	public void addAttribute(Attribute attribute) {
		if (attributes == null)
			attributes = new ArrayList<>();
		attributes.add(attribute);
	}

	public void removeAttribute(Attribute attribute) {
		if (attributes == null || attributes.isEmpty())
			return;
		if (attributes.contains(attribute))
			attributes.remove(attribute);
	}

	public List<String> getAttrValues() {
		return attrValues;
	}

	public void setAttrValues(List<String> attrValues) {
		this.attrValues = attrValues;
	}

	public void addAttrValue(String attrValue) {
		if (attrValues == null)
			attrValues = new ArrayList<>();
		attrValues.add(attrValue);
	}

	public void removeAttrValue(String attrValue) {
		if (attrValues == null || attrValues.isEmpty())
			return;
		if (attrValues.contains(attrValue))
			attrValues.remove(attrValue);
	}

	@Override
	public String toString() {
		return "Model{" +
				"name='" + name + '\'' +
				", javaClassName='" + javaClassName + '\'' +
				", attributes=" + attributes +
				", attrValues=" + attrValues +
				'}';
	}
}
