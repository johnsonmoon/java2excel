package xuyihao.java2excel.util;

import org.junit.Test;
import xuyihao.java2excel.core.entity.formatted.model.annotation.Attribute;
import xuyihao.java2excel.entity.Teacher;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Arrays;

/**
 * Created by xuyh at 2018/1/3 13:38.
 */
public class AnnotationUtilsTest {
	@Test
	public void test() {
		Teacher teacher = new Teacher();
		teacher.setId("testTeacherId001");
		teacher.setName("testTeacherName001");
		teacher.setEmail("test@test.com");
		teacher.setPhoneNumber("00101010111");
		teacher.setAddress(Arrays.asList("testAddress001", "testAddress002", "testAddress003"));

		Annotation[] annotations = AnnotationUtils.getClassAnnotations(teacher.getClass());
		for (Annotation annotation : annotations) {
			System.out.println(annotation.toString());
		}

		System.out.println(AnnotationUtils.hasAnnotationModel(Teacher.class));

		for (Field field : teacher.getClass().getDeclaredFields()) {
			System.out.println(String.format("Field: [%s] has Annotation @Attribute: [%s]", field.getName(),
					AnnotationUtils.hasAnnotationAttribute(field)));
		}

		//-----------------------test-------------------------------

		if (AnnotationUtils.hasAnnotationModel(teacher.getClass()))
			System.out.println(AnnotationUtils.getAnnotationModel(teacher.getClass()).name());
		for (Field field : teacher.getClass().getDeclaredFields()) {
			if (AnnotationUtils.hasAnnotationAttribute(field)) {
				Attribute attribute = AnnotationUtils.getAnnotationAttribute(field);
				System.out.println(attribute.attrName() + " " + attribute.attrType() + " "
						+ attribute.formatInfo()
						+ " " + attribute.defaultValue()
						+ " " + attribute.unit());

			}
		}
	}
}
