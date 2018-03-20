package xuyihao.java2excel.core;

import org.junit.Assert;
import org.junit.Test;
import xuyihao.java2excel.core.entity.free.CellData;

/**
 * Created by xuyh at 2018/3/20 17:23.
 */
public class FreeReaderTest {
	@Test
	public void test() {
		FreeReader reader = new FreeReader("C:\\Users\\Johnson\\Desktop\\FreeReader.xlsx");
		CellData cellData = reader.readExcelCellData(0, 10, 0);
		Assert.assertNotNull(cellData);
		System.out.println(cellData);
		String str = reader.readExcelData(0, 10, 0);
		Assert.assertNotNull(str);
		System.out.println(str);
		reader.close();
	}
}
