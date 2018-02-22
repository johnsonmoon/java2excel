package xuyihao.java2excel.entity;

import java.util.Date;

/**
 * Created by xuyh at 2018/2/12 10:59.
 */
public class Person {
	private String number;
	private String name;
	private String phone;
	private String email;
	private String address;
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
		return "Person{" +
				"number='" + number + '\'' +
				", name='" + name + '\'' +
				", phone='" + phone + '\'' +
				", email='" + email + '\'' +
				", address='" + address + '\'' +
				", time=" + time +
				'}';
	}
}
