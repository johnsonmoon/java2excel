package xuyihao.java2excel.entity;

import xuyihao.java2excel.core.entity.custom.annotation.Column;

import java.util.Date;

/**
 * Created by xuyh at 2018/2/12 11:00.
 */
public class PersonA {
	@Column(column = 0)
	private String number;
	@Column(column = 1)
	private String name;
	@Column(column = 2)
	private String phone;
	@Column(column = 3)
	private String email;
	private String address;
	@Column(column = 4)
	private Date time;

	public String getNumber() {
		return number;
	}

	public void setNumber(String number) {
		this.number = number;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public Date getTime() {
		return time;
	}

	public void setTime(Date time) {
		this.time = time;
	}

	@Override
	public String toString() {
		return "PersonA{" +
				"number='" + number + '\'' +
				", name='" + name + '\'' +
				", phone='" + phone + '\'' +
				", email='" + email + '\'' +
				", address='" + address + '\'' +
				", time=" + time +
				'}';
	}
}
