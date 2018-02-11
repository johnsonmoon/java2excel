package xuyihao.java2excel.core.operation;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import xuyihao.java2excel.core.operation.formatted.FormattedImporter;

/**
 * Created by xuyh at 2017/12/28 17:49.
 */
public class FormattedImporterTest {
	@Test
	public void test() throws Exception {
		String filePath = "C:\\Users\\johnson\\Desktop\\test.xlsx";
		Workbook workbook = new XSSFWorkbook(filePath);

		System.out.println("Model:");
		System.out.println(FormattedImporter.getModelFromExcel(workbook, 0));

		System.out.println("\n\nDataCount");
		int dataCount = FormattedImporter.getDataCountFromExcel(workbook, 0);
		System.out.println(dataCount);

		System.out.println("\n\nData");
		FormattedImporter.getDataFromExcel(workbook, 0, 6, dataCount).forEach(System.out::println);
		Common.closeWorkbook(workbook);
	}
}
