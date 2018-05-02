package com.github.johnsonmoon.java2excel.entity;

import java.util.Date;
import java.util.List;

/**
 * Created by xuyh at 2017/12/30 16:44.
 */
@SuppressWarnings("all")
public class Student {
	public static final String COLLECTION_NAME = "Student";
	public static String FIELD_CODE_NAME = "name";
	private static String KIND = "UNKNOWN";
	private String name;
	private String number;
	private String phoneNumber;
	private String email;
	private int score;
	private double avgScore;
	private List<String> addresses;
	private Date date;

	public Student() {
	}

	public Student(String name, String number, String phoneNumber, String email) {
		this.name = name;
		this.number = number;
		this.phoneNumber = phoneNumber;
		this.email = email;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getNumber() {
		return number;
	}

	public void setNumber(String number) {
		this.number = number;
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

	public List<String> getAddresses() {
		return addresses;
	}

	public void setAddresses(List<String> addresses) {
		this.addresses = addresses;
	}

	public int getScore() {
		return score;
	}

	public void setScore(int score) {
		this.score = score;
	}

	public double getAvgScore() {
		return avgScore;
	}

	public void setAvgScore(double avgScore) {
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
		return "Student [name=" + name + ", number=" + number + ", phoneNumber=" + phoneNumber + ", email=" + email
				+ ", score=" + score + ", avgScore=" + avgScore + ", addresses=" + addresses + "]";
	}

}
