package xuyihao.java2excel;

import org.junit.Assert;
import org.junit.Test;
import xuyihao.java2excel.core.entity.model.Model;
import xuyihao.java2excel.entity.Student;
import xuyihao.java2excel.util.JsonUtils;

import java.util.List;

/**
 * Created by xuyh at 2018/1/8 14:25.
 */
public class ReaderTest {
	@Test
	public void testModelData() {
		Reader reader = new Reader("D:\\complex.xlsx");
		Model model = reader.readExcelModel(0);
		Assert.assertNotNull(model);
		System.out.println(JsonUtils.obj2JsonStr(model));
		int count = reader.readExcelDataCount(0);
		Assert.assertNotEquals(0, count);
		System.out.println(count);

		System.out.println("\n----------------------------------------------------------------------read list");
		int listCount;
		do {
			List<Model> models = reader.readExcelData(0, 10);
			listCount = models.size();
			models.forEach(model1 -> System.out.println(JsonUtils.obj2JsonStr(model1)));
		} while (listCount > 0);

		reader.refresh();

		System.out.println("\n----------------------------------------------------------------------read array");
		Model[] models = new Model[10];
		int arrayCount;
		while ((arrayCount = reader.readExcelData(0, models)) > 0) {
			for (int i = 0; i < arrayCount; i++) {
				System.out.println(JsonUtils.obj2JsonStr(models[i]));
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
