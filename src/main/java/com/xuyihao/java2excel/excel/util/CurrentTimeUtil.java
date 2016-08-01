package com.xuyihao.java2excel.excel.util;

import java.util.Calendar;

/**
 * 获取当前系统时间的工具
 *
 * Created by Xuyh at 2016/08/01 下午 05:16.
 */
public class CurrentTimeUtil {
    public static String getCurrentTime(){
        String time = "";
        Calendar calendar = Calendar.getInstance();
        int year = calendar.get(Calendar.YEAR);
        int month = calendar.get(Calendar.MONTH);
        int date = calendar.get(Calendar.DATE);
        int hour = calendar.get(Calendar.HOUR_OF_DAY);
        int minute = calendar.get(Calendar.MINUTE);
        int second = calendar.get(Calendar.SECOND);
        time = year + "-" + month + "-" + date + "--" + hour + "-" + minute + "-" + second;
        return time;
    }
}
