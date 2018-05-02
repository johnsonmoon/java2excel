package com.github.johnsonmoon.java2excel.core;

import com.github.johnsonmoon.java2excel.util.JsonUtils;
import org.junit.Assert;
import org.junit.Test;
import com.github.johnsonmoon.java2excel.core.FormattedReader;
import com.github.johnsonmoon.java2excel.core.entity.formatted.model.Model;
import com.github.johnsonmoon.java2excel.entity.Student;

import java.util.List;

/**
 * Created by xuyh at 2018/1/8 14:25.
 */
public class FormattedReaderTest {
	@Test
	public void testModelData() {
		FormattedReader formattedReader = new FormattedReader("D:\\complex.xlsx");
		Model model = formattedReader.readExcelModel(0);
		Assert.assertNotNull(model);
		System.out.println(JsonUtils.obj2JsonStr(model));
		int count = formattedReader.readExcelDataCount(0);
		Assert.assertNotEquals(0, count);
		System.out.println(count);

		System.out.println("\n----------------------------------------------------------------------read list");
		int listCount;
		do {
			List<Model> models = formattedReader.readExcelData(0, 10);
			listCount = models.size();
			models.forEach(model1 -> System.out.println(JsonUtils.obj2JsonStr(model1)));
		} while (listCount > 0);

		formattedReader.refresh();

		System.out.println("\n----------------------------------------------------------------------read array");
		Model[] models = new Model[10];
		int arrayCount;
		while ((arrayCount = formattedReader.readExcelData(0, models)) > 0) {
			for (int i = 0; i < arrayCount; i++) {
				System.out.println(JsonUtils.obj2JsonStr(models[i]));
			}
		}

		formattedReader.close();
	}

	@Test
	public void testTypeData() {
		FormattedReader formattedReader = new FormattedReader("D:\\complex.xlsx");
		Class<?> clazz = formattedReader.readExcelJavaClass(1);
		Assert.assertNotNull(clazz);
		System.out.println(JsonUtils.obj2JsonStr(clazz));
		int count = formattedReader.readExcelDataCount(1);
		Assert.assertNotEquals(1, count);
		System.out.println(count);

		System.out.println("\n----------------------------------------------------------------------read list");
		int listCount;
		do {
			List<Student> studentList = formattedReader.readExcelData(1, 10, Student.class);
			listCount = studentList.size();
			studentList.forEach(student -> System.out.println(JsonUtils.obj2JsonStr(student)));
		} while (listCount > 0);

		formattedReader.refresh();

		System.out.println("\n----------------------------------------------------------------------read array");
		Student[] students = new Student[10];
		int arrayCount;
		while ((arrayCount = formattedReader.readExcelData(1, students, Student.class)) > 0) {
			for (int i = 0; i < arrayCount; i++) {
				System.out.println(JsonUtils.obj2JsonStr(students[i]));
			}
		}

		formattedReader.close();
	}
}
