package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import xuyihao.java2excel.core.entity.custom.map.ColumnMapper;
import xuyihao.java2excel.entity.Person;
import xuyihao.java2excel.entity.PersonA;
import xuyihao.java2excel.util.DateUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018年2月12日 上午10:55:28.
 */
public class CustomAbstractEditorTest {
	private String fileDir = System.getProperty("user.dir") + File.separator + "src/test/resources";
	private CustomAbstractEditor customAbstractEditor = new CustomAbstractEditor() {
	};
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
	public void testGenerateColumnMapper() {
		ColumnMapper columnMapper = customAbstractEditor.generateColumnMapper(Person.class);
		Assert.assertNotNull(columnMapper);
		System.out.println(columnMapper);
		ColumnMapper columnMapper1 = customAbstractEditor.generateColumnMapper(PersonA.class);
		Assert.assertNotNull(columnMapper1);
		System.out.println(columnMapper1);
	}

	@Test
	public void testGenerateData() {
		ColumnMapper columnMapper = customAbstractEditor.generateColumnMapper(Person.class);
		List<Map<String, Object>> dataMapList = customAbstractEditor.generateData(people, columnMapper);
		Assert.assertNotNull(dataMapList);
		System.out.println(dataMapList);
	}

	@Test
	public void testWriteColumnMapperHeaderColumnMapperWorkbookIntString() throws Exception {
		String filePathName = fileDir + "/testHeader1.xlsx";
		ColumnMapper columnMapper = customAbstractEditor.generateColumnMapper(Person.class);
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(customAbstractEditor.writeColumnMapperHeader(columnMapper, workbook, 0, "testSheet"));
		Assert.assertTrue(customAbstractEditor.flush(workbook, filePathName));
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testWriteColumnMapperHeaderClassOfQWorkbookIntString() throws Exception {
		String filePathName = fileDir + "/testHeader2.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(customAbstractEditor.writeColumnMapperHeader(PersonA.class, workbook, 0, "testSheet2"));
		Assert.assertTrue(customAbstractEditor.flush(workbook, filePathName));
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testWriteDataDirectlyListOfMapOfIntegerObjectWorkbookIntInt() throws Exception {
		String filePathName = fileDir + "/testDataDirectlyIntObj.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(customAbstractEditor.writeDataDirectly(intDataList, workbook, 0, 0));
		Assert.assertTrue(customAbstractEditor.flush(workbook, filePathName));
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testWriteDataDirectlyListOfMapOfStringObjectColumnMapperWorkbookIntInt() throws Exception {
		String filePathName = fileDir + "/testDataDirectlyStrObj.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(customAbstractEditor.writeDataDirectly(strDataList, columnMapper, workbook, 0, 0));
		Assert.assertTrue(customAbstractEditor.flush(workbook, filePathName));
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testWriteDataColumnMapper() throws Exception {
		String filePathName = fileDir + "/testDataColumnMapper.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(customAbstractEditor.writeData(personAS, columnMapper, workbook, 0, 0));
		Assert.assertTrue(customAbstractEditor.flush(workbook, filePathName));
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testWriteData() throws Exception {
		String filePathName = fileDir + "/testData.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(customAbstractEditor.writeData(people, workbook, 0, 0));
		Assert.assertTrue(customAbstractEditor.flush(workbook, filePathName));
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testReadDataDirectlyWorkbookIntIntInt() throws Exception {
		String filePathName = fileDir + "/testData.xlsx";
		Workbook workbook = new XSSFWorkbook(filePathName);
		List<Map<Integer, Object>> mapList = customAbstractEditor.readDataDirectly(workbook, 0, 0, 10);
		Assert.assertNotNull(mapList);
		System.out.println(mapList);
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testReadDataDirectlyColumnMapperWorkbookIntIntInt() throws Exception {
		String filePathName = fileDir + "/testDataColumnMapper.xlsx";
		Workbook workbook = new XSSFWorkbook(filePathName);
		List<Map<String, Object>> mapList = customAbstractEditor.readDataDirectly(columnMapper, workbook, 0, 0, 10);
		Assert.assertNotNull(mapList);
		System.out.println(mapList);
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testReadDataClassOfTColumnMapperWorkbookIntIntInt() throws Exception {
		String filePathName = fileDir + "/testDataColumnMapper.xlsx";
		Workbook workbook = new XSSFWorkbook(filePathName);
		List<PersonA> personAS = customAbstractEditor.readData(PersonA.class, columnMapper, workbook, 0, 0, 10);
		Assert.assertNotNull(personAS);
		System.out.println(personAS);
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testReadDataClassOfTWorkbookIntIntInt() throws Exception {
		String filePathName = fileDir + "/testData.xlsx";
		Workbook workbook = new XSSFWorkbook(filePathName);
		List<Person> people = customAbstractEditor.readData(Person.class, workbook, 0, 0, 10);
		Assert.assertNotNull(people);
		System.out.println(people);
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}

	@Test
	public void testReadDataCount() throws Exception {
		String filePathName = fileDir + "/testData.xlsx";
		Workbook workbook = new XSSFWorkbook(filePathName);
		long count = customAbstractEditor.readDataCount(workbook, 0, 0);
		Assert.assertNotEquals(0, count);
		System.out.println(count);
		Assert.assertTrue(customAbstractEditor.close(workbook));
	}
}
