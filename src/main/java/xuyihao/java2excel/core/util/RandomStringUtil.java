package xuyihao.java2excel.core.util;

import java.util.Random;

/**
 * Created by Xuyh at 2016/07/28 下午 07:47.
 */
public class RandomStringUtil {
    /**
     * 生成指定长度的随机字符串
     *
     * @param length 生成字符串的长度
     * @return 随机字符串
     */
    public static String getRandomString(int length) {
        String base = "201607281948abcdhuiqyadghiuygfc";
        Random random = new Random();
        StringBuffer sb = new StringBuffer();
        for (int i = 0; i < length; i++) {
            int number = random.nextInt(base.length());
            sb.append(base.charAt(number));
        }
        return sb.toString();
    }
}
