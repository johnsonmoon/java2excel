package com.xuyihao.java2excel.excel.util;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Xuyh at 2016/08/08 下午 03:16.
 */
public class TimeUtil {
    /**
     * 一天二十四小时的毫秒数
     *
     */
    public static long ONE_DAY_TIME_MILLISECONDS = 24 * 60 * 60 * 1000;
    /**
     * 半天十二小时的毫秒数
     *
     */
    public static long HALF_DAY_TIME_MILLISECONDS = 12 * 60 * 60 * 1000;

    /**
     * 获取指定时间的毫秒数
     *
     * @param time 格式"HH:MM:SS"
     * @return
     */
    public static long getTimeMillis(String time) {
        try {
            DateFormat dateFormat = new SimpleDateFormat("yy-MM-dd HH:mm:ss");
            DateFormat dayFormat = new SimpleDateFormat("yy-MM-dd");
            Date curDate = dateFormat.parse(dayFormat.format(new Date()) + " " + time);
            return curDate.getTime();
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return 0;
    }

    /**
     * 获取输入目标时间距离当前时间的所剩时间距离
     *
     * @param targetTime 目标时间，格式("HH:MM:SS")
     * @return 时间差(Unit: MILLISECONDS)
     */
    public static long getRemainedTimeMillis(String targetTime) {
        long time = getTimeMillis(targetTime);
        long returnTime = time - System.currentTimeMillis();
        if (returnTime < 0) {
            returnTime += ONE_DAY_TIME_MILLISECONDS;
        }
        return returnTime;
    }

    /**
     * 获取当前时间
     *
     * @param timeFormat 时间格式（如："yyyy-MM-dd-HH-mm-ss"）
     * @return
     */
    public static String getCurrentTime(String timeFormat) {
        Date dt = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat(timeFormat);
        String currentTime = sdf.format(dt);
        return currentTime;
    }
}
