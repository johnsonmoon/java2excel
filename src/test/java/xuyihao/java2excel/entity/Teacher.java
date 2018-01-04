package xuyihao.java2excel.entity;

import xuyihao.java2excel.core.entity.annotation.Attribute;
import xuyihao.java2excel.core.entity.annotation.Template;

import java.util.List;

/**
 * Created by xuyh at 2018/1/3 13:41.
 */
@Template(name = "Teacher")
public class Teacher {
	public static final String COLLLECTION_NAME = "Teacher";
	public static final String FIELD_CODE_ID = "id";
	public static final String FIELD_CODE_NAME = "name";
	public static final String FIELD_CODE_PHONE_NUMBER = "phoneNumber";
	public static final String FIELD_CODE_EMAIL = "email";
	public static final String FIELD_CODE_ADDRESS = "address";

	@Attribute(attrCode = "id", attrName = "ID", attrType = "String", formatInfo = "text")
	private String id;
	@Attribute(attrCode = "name", attrName = "NAME", attrType = "String", formatInfo = "text")
	private String name;
	@Attribute(attrCode = "phoneNumber", attrName = "PHONE_NUMBER", attrType = "String", formatInfo = "text")
	private String phoneNumber;
	@Attribute(attrCode = "email", attrName = "EMAIL", attrType = "String", formatInfo = "text")
	private String email;
	@Attribute(attrCode = "address", attrName = "ADDRESS", attrType = "List<String>", formatInfo = "[\"a\", \"b\", ...]")
	private List<String> address;
	@Attribute(attrCode = "avgScore", attrName = "AVG_SCORE", attrType = "Integer", formatInfo = "text")
	private int avgScore;

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

	@Override
	public String toString() {
		return "Teacher{" +
				"id='" + id + '\'' +
				", name='" + name + '\'' +
				", phoneNumber='" + phoneNumber + '\'' +
				", email='" + email + '\'' +
				", address=" + address +
				", avgScore=" + avgScore +
				'}';
	}
}
