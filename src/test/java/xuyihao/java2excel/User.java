package xuyihao.java2excel;

import java.util.Date;

/**
 * Created by xuyh at 2018/2/22 0022 下午 18:03.
 */
public class User {
	private int number;
	private String name;
	private Date registerTime;

	public int getNumber() {
		return number;
	}

	public void setNumber(int number) {
		this.number = number;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Date getRegisterTime() {
		return registerTime;
	}

	public void setRegisterTime(Date registerTime) {
		this.registerTime = registerTime;
	}

	@Override
	public String toString() {
		return "User{" +
				"number=" + number +
				", name='" + name + '\'' +
				", registerTime=" + registerTime +
				'}';
	}
}
