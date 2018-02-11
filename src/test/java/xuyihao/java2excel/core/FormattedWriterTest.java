package xuyihao.java2excel.core;

import org.junit.Assert;
import org.junit.Test;
import xuyihao.java2excel.core.FormattedWriter;
import xuyihao.java2excel.entity.Student;
import xuyihao.java2excel.entity.Teacher;
import xuyihao.java2excel.util.DateUtils;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by xuyh at 2018/1/5 18:03.
 */
public class FormattedWriterTest {
	@Test
	public void test() {
		List<Teacher> teachers = new ArrayList<>();
		for (int i = 0; i < 100; i++) {
			Teacher teacher = new Teacher();
			teacher.setId("testId" + i);
			teacher.setName("testName" + i);
			teacher.setEmail("testEmail@xxx.com");
			teacher.setPhoneNumber("111111111" + i);
			teacher.setAddress(Arrays.asList("testAddress1" + i, "testAddress2" + i, "testAddress3" + i));
			teacher.setDate(DateUtils.currentDateTimeForDate());
			teachers.add(teacher);
		}
		List<Student> students = new ArrayList<>();
		for (int i = 0; i < 100; i++) {
			Student student = new Student();
			student.setName("name" + (i + 1));
			student.setEmail("email" + (i + 1));
			student.setNumber("number" + (i + 1));
			student.setPhoneNumber("phoneNumber" + (i + 1));
			student.setAddresses(Arrays.asList("address0" + (i + 1), "address1" + (i + 1)));
			student.setDate(DateUtils.currentDateTimeForDate());
			students.add(student);
		}

		FormattedWriter formattedWriter = new FormattedWriter("D:\\complex.xlsx");

		formattedWriter.setLanguage("en_US");

		Assert.assertTrue(formattedWriter.writeExcelModel(Teacher.class, 0));
		Assert.assertTrue(formattedWriter.writeExcelModel(Student.class, 1));

		for (int i = 0; i < 100; i += 10) {

			List<Teacher> t = new ArrayList<>();
			for (int j = i; j < i + 10; j++) {
				t.add(teachers.get(j));
			}
			Assert.assertTrue(formattedWriter.writeExcelData(t, 0));

			List<Student> s = new ArrayList<>();
			for (int k = i; k < i + 10; k++) {
				s.add(students.get(k));
			}
			Assert.assertTrue(formattedWriter.writeExcelData(s, 1));

		}

		Assert.assertTrue(formattedWriter.flush());
		Assert.assertTrue(formattedWriter.close());
	}
}
