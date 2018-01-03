package xuyihao.java2excel.util;

import org.junit.Test;
import xuyihao.java2excel.entity.Student;

/**
 * Created by xuyh at 2018/1/3 11:34.
 */
public class ReflectionUtilsTest {
	@Test
	public void testSetFieldValue() {
		Student student = new Student();
		student.setName("Johnson");

		System.out.println(ReflectionUtils.getFieldValue(student, "name"));
		System.out.println(ReflectionUtils.setFieldValue(student, "name", "John"));
		System.out.println(student.getName());
	}
}
