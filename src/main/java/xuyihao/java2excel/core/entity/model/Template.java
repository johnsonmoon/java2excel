package xuyihao.java2excel.core.entity.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 数据模型
 * <p>
 * Created by Xuyh at 2016/07/21 下午 01:40.
 */
public class Template {
	private String name;//模型名称(Sheet名称)
	private List<Attribute> attributes;//模型的属性
	private List<String> attrValues;//模型的属性值

	public Template() {
	}

	public Template(String name, List<Attribute> attributes, List<String> attrValues) {
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
		return "Template{" +
				"name='" + name + '\'' +
				", attributes=" + attributes +
				", attrValues=" + attrValues +
				'}';
	}
}
