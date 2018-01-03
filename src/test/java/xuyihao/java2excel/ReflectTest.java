package xuyihao.java2excel;

import org.junit.Test;
import xuyihao.java2excel.entity.Student;

import java.lang.reflect.Field;

/**
 * Created by xuyh at 2017/12/30 16:43.
 */
public class ReflectTest {
	@Test
	public void test() {
		Class<?> tClass = Student.class;
		System.out.println(String.format("className: [%s]", tClass.getName()));
		Field[] fields = tClass.getDeclaredFields();
		for (Field field : fields) {
			System.out.println("\n\n\nField-----------------------------");
			System.out.println(String.format("toString:[%s]", field.toString()));
			System.out.println(field.getType());

			String typeStr = field.getType().toString();
			System.out.println(typeStr);
			String lessTypeStr = typeStr.substring(typeStr.lastIndexOf(".") + 1, typeStr.length());
			System.out.println(lessTypeStr);
		}
	}
}
