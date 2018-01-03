package xuyihao.java2excel;

import org.junit.Ignore;
import org.junit.Test;
import xuyihao.java2excel.entity.Student;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by xuyh at 2018/1/2 14:00.
 */
public class WriterTest {
	@Ignore
	public void testExcelTemplate() {
		Writer writer = new Writer();
		writer.setTemplateLanguage("zh_CN");
		System.out.println(writer.writeExcelTemplate(Student.class, "D:\\test.xlsx"));
	}

	@Test
	public void testExcelData() {
		Writer writer = new Writer();
		writer.setTemplateLanguage("zh_CN");

		List<Student> students = new ArrayList<>();
		for (int i = 0; i < 10; i++) {
			Student student = new Student();
			student.setName("name" + (i + 1));
			student.setEmail("email" + (i + 1));
			student.setNumber("number" + (i + 1));
			student.setPhoneNumber("phoneNumber" + (i + 1));
			student.setAddresses(Arrays.asList("address0" + (i + 1), "address1" + (i + 1)));
			students.add(student);
		}

		System.out.println(writer.writeExcelData(students, "D:\\test.xlsx"));
	}
}
