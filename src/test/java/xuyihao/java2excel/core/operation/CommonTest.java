package xuyihao.java2excel.core.operation;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;
import xuyihao.java2excel.core.entity.common.CellStyle;
import xuyihao.java2excel.util.FileUtils;

/**
 * Created by xuyh at 2018/3/21 10:50.
 */
public class CommonTest {
	@Test
	public void test() throws Exception {
		String filePathName = "C:\\Users\\Johnson\\Desktop\\Common-test.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = Common.getSheetCreateIfNotExist(workbook, 0, "test");
		Assert.assertTrue(Common.setColumnWidth(sheet, 0, 6000));
		Assert.assertTrue(Common.hideColumn(sheet, 10));
		for (int row = 0; row < 20; row++) {
			for (int column = 0; column < 20; column++) {
				Common.insertCellValue(sheet, column, row, "test_" + row + "_" + column,
						Common.createCellStyle(workbook, CellStyle.CELL_STYLE_TYPE_VALUE.getCode()));
			}
		}
		Assert.assertTrue(Common.setRowHeight(sheet, 0, 300));
		Assert.assertTrue(Common.hideRow(sheet, 5));
		Assert.assertTrue(Common.writeFileToDisk(workbook, FileUtils.createFileDE(filePathName)));
		Assert.assertTrue(Common.closeWorkbook(workbook));
	}

	@Test
	public void rowHeightTest() throws Exception {
		String filePathName = "C:\\Users\\Johnson\\Desktop\\Common-test-row-height.xlsx";
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = Common.getSheetCreateIfNotExist(workbook, 0, "test");
		Assert.assertTrue(Common.setDefaultRowHeight(sheet, 600));
		for (int row = 0; row < 10; row++) {
			for (int column = 0; column < 10; column++) {
				Common.insertCellValue(sheet, column, row, "test_" + row + "_" + column,
						Common.createCellStyle(workbook, CellStyle.CELL_STYLE_TYPE_VALUE.getCode()));
			}
		}
		Assert.assertTrue(Common.setRowHeight(sheet, 1, 600));
		Assert.assertTrue(Common.writeFileToDisk(workbook, FileUtils.createFileDE(filePathName)));
		Assert.assertTrue(Common.closeWorkbook(workbook));
	}
}
