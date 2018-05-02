package com.github.johnsonmoon.java2excel.core.operation;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import com.github.johnsonmoon.java2excel.core.entity.custom.ColumnMapper;
import com.github.johnsonmoon.java2excel.core.operation.custom.CustomExporter;
import com.github.johnsonmoon.java2excel.util.FileUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/2/11 16:48.
 */
public class CustomExporterTest {
	private String fileDir = System.getProperty("user.dir") + File.separator + "src/test/resources";
	private ColumnMapper columnMapper = new ColumnMapper();
	private List<Map<String, Object>> strDataList = new ArrayList<>();
	private List<Map<Integer, Object>> intDataList = new ArrayList<>();

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
		}
	}

	@Test
	public void testCreateExcel() throws Exception {
		String filePathName = fileDir + "/testCreateExcel.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(CustomExporter.createExcel(workbook, 0, "testSheet", columnMapper));
		Assert.assertTrue(Common.writeFileToDisk(workbook, FileUtils.createFileDE(filePathName)));
	}

	@Test
	public void testInsertDataIntegerValue() throws Exception {
		String filePathName = fileDir + "/testInsertDataIntegerValue.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(CustomExporter.insertExcelData(workbook, 0, 0, intDataList));
		Assert.assertTrue(Common.writeFileToDisk(workbook, FileUtils.createFileDE(filePathName)));
	}

	@Test
	public void testInsertDataStringValue() throws Exception {
		String filePathName = fileDir + "/testInsertDataStringValue.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Assert.assertTrue(CustomExporter.insertExcelData(workbook, 0, 0, columnMapper, strDataList));
		Assert.assertTrue(Common.writeFileToDisk(workbook, FileUtils.createFileDE(filePathName)));
	}
}
