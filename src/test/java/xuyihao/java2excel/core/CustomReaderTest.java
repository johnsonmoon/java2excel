package xuyihao.java2excel.core;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import xuyihao.java2excel.entity.Person;
import xuyihao.java2excel.util.DateUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by xuyh at 2018/2/12 15:33.
 */
public class CustomReaderTest {
	private String filePathName = System.getProperty("user.dir") + File.separator
			+ "src/test/resources/testCustomReader.xlsx";
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
		CustomWriter writer = new CustomWriter(filePathName);
		Assert.assertTrue(writer.writeExcelMetaInfo(Person.class, 0));
		Assert.assertTrue(writer.writeExcelData(people, 0));
		Assert.assertTrue(writer.flush());
		Assert.assertTrue(writer.close());
	}

	@Test
	public void test() {
		CustomReader reader = new CustomReader(filePathName, 1);
		long count = reader.readExcelDataCount(0);
		Assert.assertNotEquals(0, count);
		System.out.println(count);
		List<Person> personList;
		do {
			personList = reader.readExcelData(0, 10, Person.class);
			personList.forEach(System.out::println);
		} while (!personList.isEmpty());

		Assert.assertTrue(reader.refresh());

		System.out.println("\r\r-----------------------------\r\r");

		Person[] personArray = new Person[10];
		int readLength;
		while ((readLength = reader.readExcelData(0, personArray, Person.class)) > 0) {
			for (int i = 0; i < readLength; i++) {
				System.out.println(personArray[i]);
			}
		}

		Assert.assertTrue(reader.close());
	}
}
