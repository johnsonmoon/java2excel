package xuyihao.java2excel.core;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import xuyihao.java2excel.core.entity.custom.ColumnMapper;
import xuyihao.java2excel.entity.Person;
import xuyihao.java2excel.entity.PersonA;
import xuyihao.java2excel.util.DateUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/2/12 13:47.
 */
public class CustomEditorTest {
	private String fileDir = System.getProperty("user.dir") + File.separator + "src/test/resources";
	private ColumnMapper columnMapper = new ColumnMapper();
	private List<Map<String, Object>> strDataList = new ArrayList<>();
	private List<Map<Integer, Object>> intDataList = new ArrayList<>();
	private List<Person> people = new ArrayList<>();
	private List<PersonA> personAS = new ArrayList<>();

	@Before
	public void setUp() {
		columnMapper.add(0, "number");
		columnMapper.add(1, "name");
		columnMapper.add(2, "phone");
		columnMapper.add(3, "email");
		columnMapper.add(4, "address");

		for (int i = 0; i < 20; i++) {
			Map<String, Object> strData = new HashMap<>();
			strData.put("number", "number_00" + i);
			strData.put("name", "name_00" + i);
			strData.put("phone", "phone_00" + i);
			strData.put("email", "email_00" + i);
			strData.put("address", "address00_" + i);
			strDataList.add(strData);
			Map<Integer, Object> intData = new HashMap<>();
			intData.put(0, "number_00" + i);
			intData.put(1, "name_00" + i);
			intData.put(2, "phone_00" + i);
			intData.put(3, "email_00" + i);
			intData.put(4, "address00_" + i);
			intDataList.add(intData);
			Person person = new Person();
			person.setNumber("number_00" + i);
			person.setName("name_00" + i);
			person.setPhone("phone_00" + i);
			person.setEmail("email_00" + i);
			person.setAddress("address00_" + i);
			person.setTime(DateUtils.currentDateTimeForDate());
			people.add(person);
			PersonA personA = new PersonA();
			personA.setNumber("number_00" + i);
			personA.setName("name_00" + i);
			personA.setPhone("phone_00" + i);
			personA.setEmail("email_00" + i);
			personA.setAddress("address00_" + i);
			personA.setTime(DateUtils.currentDateTimeForDate());
			personAS.add(personA);
		}
	}

	@Test
	public void testCase1() {
		String filePathName = fileDir + "/testCustomEditor1.xlsx";
		String filePathName2 = fileDir + "/testCustomEditor2.xlsx";

		CustomEditor customEditor = new CustomEditor(filePathName);
		Assert.assertTrue(customEditor.writeExcelMetaInfo(columnMapper, 0, "TestSheet"));
		Assert.assertTrue(customEditor.writeExcelDataDirectly(intDataList, 0, 1));
		Assert.assertTrue(customEditor.flush());
		Assert.assertTrue(customEditor.close());

		CustomEditor customEditor1 = new CustomEditor(filePathName);
		Assert.assertTrue(customEditor1.writeExcelMetaInfo(columnMapper, 1, "testSheet2"));
		Assert.assertTrue(customEditor1.writeExcelDataDirectly(strDataList, columnMapper, 1, 1));
		Assert.assertTrue(customEditor1.flush(filePathName2));//save into another file
		Assert.assertTrue(customEditor1.close());

		CustomEditor customEditor2 = new CustomEditor(filePathName2);
		long count = customEditor2.readExcelDataCount(0, 1);
		Assert.assertNotEquals(0, count);
		System.out.println(count);
		List<Map<Integer, Object>> mapList = customEditor2.readExcelDataDirectly(0, 1, 10);
		Assert.assertNotNull(mapList);
		System.out.println(mapList);
		List<Map<String, Object>> mapList1 = customEditor2.readExcelDataDirectly(columnMapper, 1, 1, 10);
		Assert.assertNotNull(mapList1);
		System.out.println(mapList1);
		Assert.assertTrue(customEditor2.close());
	}

	@Test
	public void testCase2() {
		String filePathName3 = fileDir + "/testCustomEditor3.xlsx";
		String filePathName4 = fileDir + "/testCustomEditor4.xlsx";

		CustomEditor customEditor = new CustomEditor(filePathName3);
		Assert.assertTrue(customEditor.writeExcelMetaInfo(Person.class, 0, "Person"));
		Assert.assertTrue(customEditor.writeExcelData(people, columnMapper, 0, 1));
		Assert.assertTrue(customEditor.flush());
		Assert.assertTrue(customEditor.close());

		CustomEditor customEditor1 = new CustomEditor(filePathName3);
		Assert.assertTrue(customEditor1.writeExcelMetaInfo(PersonA.class, 1, "PersonA"));
		Assert.assertTrue(customEditor1.writeExcelData(personAS, 1, 1));
		Assert.assertTrue(customEditor1.flush(filePathName4));//save into another file
		Assert.assertTrue(customEditor1.close());

		CustomEditor customEditor2 = new CustomEditor(filePathName4);
		long count = customEditor2.readExcelDataCount(0, 1);
		Assert.assertNotEquals(0, count);
		System.out.println(count);
		List<Person> personList = customEditor2.readExcelData(Person.class, columnMapper, 0, 1, 10);
		Assert.assertNotNull(personList);
		System.out.println(personList);
		List<PersonA> personAList = customEditor2.readExcelData(1, 1, 10, PersonA.class);
		Assert.assertNotNull(personAList);
		System.out.println(personAList);
		Assert.assertTrue(customEditor2.close());
	}
}
