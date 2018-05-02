package com.github.johnsonmoon.java2excel.core;

import com.github.johnsonmoon.java2excel.util.DateUtils;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import com.github.johnsonmoon.java2excel.entity.Person;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by xuyh at 2018/2/12 15:13.
 */
public class CustomWriterTest {
	private String fileDir = System.getProperty("user.dir") + File.separator + "src/test/resources";
	private List<Person> people = new ArrayList<>();

	@Before
	public void setUp() {
		for (int i = 0; i < 20; i++) {
			Person person = new Person();
			person.setNumber("number_00" + i);
			person.setName("name_00" + i);
			person.setPhone("phone_00" + i);
			person.setEmail("email_00" + i);
			person.setAddress("address00_" + i);
			person.setTime(DateUtils.currentDateTimeForDate());
			people.add(person);
		}
	}

	@Test
	public void test() {
		CustomWriter writer = new CustomWriter(fileDir + "/testCustomWriter.xlsx");
		Assert.assertTrue(writer.writeExcelMetaInfo(Person.class, 0));
		Assert.assertTrue(writer.writeExcelData(people, 0));
		Assert.assertTrue(writer.flush());
		writer.close();
	}
}
