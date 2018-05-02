package com.github.johnsonmoon.java2excel.util;

import java.text.Format;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Xuyh at 2016/08/05 下午 06:58.
 */
public class DateUtils {
	private static final String DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";

	/**
	 * 得到当前时间
	 *
	 * @return 当前时间, 格式"yyyy-MM-dd HH:mm:ss"
	 */
	public static String currentDateTime() {
		Date dt = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat(DATE_TIME_FORMAT);
		String currentTime = sdf.format(dt);
		return currentTime;
	}

	/**
	 * 得到当前时间
	 *
	 * @return
	 */
	public static Date currentDateTimeForDate() {
		return parseDateTime(currentDateTime());
	}

	/**
	 * 得到格式化时间
	 *
	 * @return 当前时间
	 */
	public static String formatDateTime(Date date) {
		if (date == null)
			return null;
		SimpleDateFormat sdf = new SimpleDateFormat(DATE_TIME_FORMAT);
		String dateFormate = sdf.format(date);
		return dateFormate;
	}

	/**
	 * string类型转成Date
	 *
	 * @param dateTime 传入str类型
	 * @return Date
	 * @throws ParseException
	 */
	public static Date parseDateTime(String dateTime) {
		SimpleDateFormat sdf = new SimpleDateFormat(DATE_TIME_FORMAT);
		try {
			return sdf.parse(dateTime);
		} catch (ParseException e) {
			e.printStackTrace();
			return null;
		}
	}

	/**
	 * Long类型时间戳转换成Date
	 *
	 * @param timeStamp
	 * @return
	 */
	public static Date parseDateTime(Long timeStamp) {
		Format formatter = new SimpleDateFormat(DATE_TIME_FORMAT);
		return parseDateTime(formatter.format(timeStamp));
	}
}
