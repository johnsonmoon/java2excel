package xuyihao.java2excel;

import org.junit.Assert;
import org.junit.Test;
import xuyihao.java2excel.core.entity.model.Template;
import xuyihao.java2excel.entity.Student;
import xuyihao.java2excel.entity.Teacher;
import xuyihao.java2excel.util.JsonUtils;

import java.util.List;

/**
 * Created by xuyh at 2018/1/4 10:00.
 */
public class QuickReaderTest {
	@Test
	public void testReadStudentClassType() {
		QuickReader reader = new QuickReader();
		Class<?> clazz = reader.readExcelJavaClass("D:\\studentTemplate.xlsx");
		Assert.assertNotNull(clazz);
		System.out.println(clazz.getName());
	}

	@Test
	public void testReadStudentTemplate() {
		QuickReader reader = new QuickReader();
		Template template = reader.readExcelTemplate("D:\\studentTemplate.xlsx");
		Assert.assertNotNull(template);
		System.out.println(JsonUtils.obj2JsonStr(template));
	}

	@Test
	public void testReadTeacherTemplate() {
		QuickReader reader = new QuickReader();
		Template template = reader.readExcelTemplate("D:\\teacherTemplate.xlsx");
		Assert.assertNotNull(template);
		System.out.println(JsonUtils.obj2JsonStr(template));
	}

	@Test
	public void testReadStudentTemplateData() {
		QuickReader reader = new QuickReader();
		List<Template> templates = reader.readExcelDataAll("D:\\studentData.xlsx");
		Assert.assertNotNull(templates);
		System.out.println(JsonUtils.obj2JsonStr(templates));
	}

	@Test
	public void testReadTeacherTemplateData() {
		QuickReader reader = new QuickReader();
		List<Template> templates = reader.readExcelDataAll("D:\\teacherData.xlsx");
		Assert.assertNotNull(templates);
		System.out.println(JsonUtils.obj2JsonStr(templates));
	}

	@Test
	public void testReadTeacherData() {
		QuickReader reader = new QuickReader();
		List<Teacher> teachers = reader.readExcelDataAll("D:\\teacherData.xlsx", Teacher.class);
		Assert.assertNotNull(teachers);
		System.out.println(JsonUtils.obj2JsonStr(teachers));
	}

	@Test
	public void testReadStudentData() {
		QuickReader reader = new QuickReader();
		List<Student> students = reader.readExcelDataAll("D:\\studentData.xlsx", Student.class);
		Assert.assertNotNull(students);
		System.out.println(JsonUtils.obj2JsonStr(students));
	}
}
