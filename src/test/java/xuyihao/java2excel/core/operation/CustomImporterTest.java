package xuyihao.java2excel.core.operation;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xuyihao.java2excel.core.entity.map.ColumnMapper;
import xuyihao.java2excel.core.operation.custom.CustomImporter;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/2/11 15:08.
 */
public class CustomImporterTest {
	private static Logger logger = LoggerFactory.getLogger(CustomImporterTest.class);
	private String filePathName = System.getProperty("user.dir") + File.separator + "src/test/resources/test.xlsx";
	private Workbook workbook;

	@Before
	public void setUp() {
		try {
			workbook = new XSSFWorkbook(filePathName);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
		}
	}

	@Test
	public void testDataCount() {
		long count = CustomImporter.getDataCountFromExcel(workbook, 0, 1);
		Assert.assertNotEquals(0, count);
		System.out.println("Data count:  " + count);
	}

	@Test
	public void testGetDataIntegerValue() {
		List<Map<Integer, Object>> mapList = CustomImporter.getDataFromExcel(workbook, 0, 1, 25);
		System.out.println(mapList);
	}

	@Test
	public void testGetDataStringValue() {
		ColumnMapper columnMapper = new ColumnMapper();
		columnMapper.add(0, "number");
		columnMapper.add(1, "name");
		columnMapper.add(2, "phone");
		columnMapper.add(3, "email");
		columnMapper.add(4, "address");
		List<Map<String, Object>> mapList = CustomImporter.getDataFromExcel(workbook, 0, columnMapper, 1, 25);
		System.out.println(mapList);
	}
}
