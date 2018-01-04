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
public class ReaderTest {
	@Test
	public void testReadStudentClassType() {
		Reader reader = new Reader();
		Class<?> clazz = reader.readExcelJavaClass("D:\\studentTemplate.xlsx");
		Assert.assertNotNull(clazz);
		System.out.println(clazz.getName());
	}

	@Test
	public void testReadStudentTemplate() {
		Reader reader = new Reader();
		Template template = reader.readExcelTemplate("D:\\studentTemplate.xlsx");
		Assert.assertNotNull(template);
		System.out.println(JsonUtils.obj2JsonStr(template));
	}

	@Test
	public void testReadTeacherTemplate() {
		Reader reader = new Reader();
		Template template = reader.readExcelTemplate("D:\\teacherTemplate.xlsx");
		Assert.assertNotNull(template);
		System.out.println(JsonUtils.obj2JsonStr(template));
	}

	@Test
	public void testReadStudentTemplateData() {
		Reader reader = new Reader();
		List<Template> templates = reader.readExcelDataAll("D:\\studentData.xlsx");
		Assert.assertNotNull(templates);
		System.out.println(JsonUtils.obj2JsonStr(templates));
	}

	@Test
	public void testReadTeacherTemplateData() {
		Reader reader = new Reader();
		List<Template> templates = reader.readExcelDataAll("D:\\teacherData.xlsx");
		Assert.assertNotNull(templates);
		System.out.println(JsonUtils.obj2JsonStr(templates));
	}

	@Test
	public void testReadTeacherData() {
		Reader reader = new Reader();
		List<Teacher> teachers = reader.readExcelDataAll("D:\\teacherData.xlsx", Teacher.class);
		Assert.assertNotNull(teachers);
		System.out.println(JsonUtils.obj2JsonStr(teachers));
	}

	@Test
	public void testReadStudentData() {
		Reader reader = new Reader();
		List<Student> students = reader.readExcelDataAll("D:\\studentData.xlsx", Student.class);
		Assert.assertNotNull(students);
		System.out.println(JsonUtils.obj2JsonStr(students));
	}
}
