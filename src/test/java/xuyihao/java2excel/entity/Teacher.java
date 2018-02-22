package xuyihao.java2excel.entity;

import xuyihao.java2excel.core.entity.formatted.model.annotation.Attribute;
import xuyihao.java2excel.core.entity.formatted.model.annotation.Model;

import java.util.Date;
import java.util.List;

/**
 * Created by xuyh at 2018/1/3 13:41.
 */
@SuppressWarnings("all")
@Model(name = "Teacher")
public class Teacher {
	public static final String COLLLECTION_NAME = "Teacher";
	public static final String FIELD_CODE_ID = "id";
	public static final String FIELD_CODE_NAME = "name";
	public static final String FIELD_CODE_PHONE_NUMBER = "phoneNumber";
	public static final String FIELD_CODE_EMAIL = "email";
	public static final String FIELD_CODE_ADDRESS = "address";

	@Attribute(attrName = "ID", attrType = "String", formatInfo = "text")
	private String id;
	@Attribute(attrName = "NAME", attrType = "String", formatInfo = "text")
	private String name;
	@Attribute(attrName = "PHONE_NUMBER", attrType = "String", formatInfo = "text")
	private String phoneNumber;
	@Attribute(attrName = "EMAIL", attrType = "String", formatInfo = "text")
	private String email;
	@Attribute(attrName = "ADDRESS", attrType = "List<String>", formatInfo = "[\"a\", \"b\", ...]")
	private List<String> address;
	@Attribute(attrName = "AVG_SCORE", attrType = "Integer", formatInfo = "text")
	private int avgScore;
	@Attribute(attrName = "DATE", attrType = "Date", formatInfo = "yyyy-MM-dd HH:mm:ss")
	private Date date;

	public Teacher() {
	}

	public Teacher(String name, String phoneNumber, String email) {
		this.name = name;
		this.phoneNumber = phoneNumber;
		this.email = email;
	}

	private Teacher(String id, String name) {
		this.id = id;
		this.name = name;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getPhoneNumber() {
		return phoneNumber;
	}

	public void setPhoneNumber(String phoneNumber) {
		this.phoneNumber = phoneNumber;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public List<String> getAddress() {
		return address;
	}

	public void setAddress(List<String> address) {
		this.address = address;
	}

	public int getAvgScore() {
		return avgScore;
	}

	public void setAvgScore(int avgScore) {
		this.avgScore = avgScore;
	}

	public Date getDate() {
		return date;
	}

	public void setDate(Date date) {
		this.date = date;
	}

	@Override
	public String toString() {
		return "Teacher [id=" + id + ", name=" + name + ", phoneNumber=" + phoneNumber + ", email=" + email + ", address="
				+ address + ", avgScore=" + avgScore + ", date=" + date + "]";
	}

}
