package xuyihao.java2excel;

import org.junit.Assert;
import org.junit.Test;
import xuyihao.java2excel.core.entity.model.Template;
import xuyihao.java2excel.entity.Student;
import xuyihao.java2excel.util.JsonUtils;

import java.util.List;

/**
 * Created by xuyh at 2018/1/8 14:25.
 */
public class ReaderTest {
	@Test
	public void testTemplateData() {
		Reader reader = new Reader("D:\\complex.xlsx");
		Template template = reader.readExcelTemplate(0);
		Assert.assertNotNull(template);
		System.out.println(JsonUtils.obj2JsonStr(template));
		int count = reader.readExcelDataCount(0);
		Assert.assertNotEquals(0, count);
		System.out.println(count);

		System.out.println("\n----------------------------------------------------------------------read list");
		int listCount;
		do {
			List<Template> templates = reader.readExcelData(0, 10);
			listCount = templates.size();
			templates.forEach(template1 -> System.out.println(JsonUtils.obj2JsonStr(template1)));
		} while (listCount > 0);

		reader.refresh();

		System.out.println("\n----------------------------------------------------------------------read array");
		Template[] templates = new Template[10];
		int arrayCount;
		while ((arrayCount = reader.readExcelData(0, templates)) > 0) {
			for (int i = 0; i < arrayCount; i++) {
				System.out.println(JsonUtils.obj2JsonStr(templates[i]));
			}
		}

		reader.close();
	}

	@Test
	public void testTypeData() {
		Reader reader = new Reader("D:\\complex.xlsx");
		Class<?> clazz = reader.readExcelJavaClass(1);
		Assert.assertNotNull(clazz);
		System.out.println(JsonUtils.obj2JsonStr(clazz));
		int count = reader.readExcelDataCount(1);
		Assert.assertNotEquals(1, count);
		System.out.println(count);

		System.out.println("\n----------------------------------------------------------------------read list");
		int listCount;
		do {
			List<Student> studentList = reader.readExcelData(1, 10, Student.class);
			listCount = studentList.size();
			studentList.forEach(student -> System.out.println(JsonUtils.obj2JsonStr(student)));
		} while (listCount > 0);

		reader.refresh();

		System.out.println("\n----------------------------------------------------------------------read array");
		Student[] students = new Student[10];
		int arrayCount;
		while ((arrayCount = reader.readExcelData(1, students, Student.class)) > 0) {
			for (int i = 0; i < arrayCount; i++) {
				System.out.println(JsonUtils.obj2JsonStr(students[i]));
			}
		}

		reader.close();
	}
}
