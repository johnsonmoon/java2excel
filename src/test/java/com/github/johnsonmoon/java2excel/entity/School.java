package com.github.johnsonmoon.java2excel.entity;

/**
 * Created by xuyh at 2018/1/4 16:57.
 */
public class School {
	private String name;
	private String address;

	public School(String name, String address) {
		this.name = name;
		this.address = address;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}
}
