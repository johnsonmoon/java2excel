package com.github.johnsonmoon.java2excel.core;

import com.github.johnsonmoon.java2excel.util.JsonUtils;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import com.github.johnsonmoon.java2excel.core.entity.formatted.dict.Languages;
import com.github.johnsonmoon.java2excel.entity.Teacher;
import com.github.johnsonmoon.java2excel.util.FileUtils;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by xuyh at 2018/1/10 18:23.
 */
public class FormattedEditorTest {
	private static String filePathName = "D:\\teacher.xlsx";

	@Before
	public void setUp() {
		FileUtils.delete(filePathName);
		List<Teacher> teachers = new ArrayList<>();
		for (int i = 0; i < 100; i++) {
			Teacher teacher = new Teacher();
			teacher.setId("testId" + i);
			teacher.setName("testName" + i);
			teacher.setEmail("testEmail@xxx.com");
			teacher.setPhoneNumber("111111111" + i);
			teacher.setAddress(Arrays.asList("testAddress1" + i, "testAddress2" + i, "testAddress3" + i));
			teachers.add(teacher);
		}

		FormattedWriter formattedWriter = new FormattedWriter(filePathName);
		formattedWriter.setLanguage(Languages.ZH_CN.getLanguage());
		Assert.assertTrue(formattedWriter.writeExcelMetaInfo(Teacher.class, 0));
		Assert.assertTrue(formattedWriter.writeExcelData(teachers, 0));
		Assert.assertTrue(formattedWriter.flush());
		formattedWriter.close();
	}

	@Test
	public void test() {
		FormattedEditor editor = new FormattedEditor(filePathName);
		editor.setLanguage(Languages.ZH_CN.getLanguage());
		int count = editor.readExcelDataCount(0);
		System.out.println(count);
		List<Teacher> readTeachers = editor.readExcelData(0, 0, 10, Teacher.class);
		readTeachers.forEach(teacher -> System.out.println(JsonUtils.obj2JsonStr(teacher)));

		List<Teacher> teachers = new ArrayList<>();
		for (int i = 100; i < 120; i++) {
			Teacher teacher = new Teacher();
			teacher.setId("testId" + i);
			teacher.setName("testName" + i);
			teacher.setEmail("testEmail@xxx.com");
			teacher.setPhoneNumber("111111111" + i);
			teacher.setAddress(Arrays.asList("testAddress1" + i, "testAddress2" + i, "testAddress3" + i));
			teachers.add(teacher);
		}
		Assert.assertTrue(editor.writeExcelData(teachers, 0, count + 10));
		Assert.assertTrue(editor.flush("D:\\test2.xlsx"));
		editor.close();
	}
}
