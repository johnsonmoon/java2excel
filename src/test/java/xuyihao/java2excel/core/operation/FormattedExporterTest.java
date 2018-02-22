package xuyihao.java2excel.core.operation;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import xuyihao.java2excel.core.entity.formatted.model.Attribute;
import xuyihao.java2excel.core.entity.formatted.model.Model;
import xuyihao.java2excel.core.operation.formatted.FormattedExporter;
import xuyihao.java2excel.util.FileUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by xuyh at 2017/12/28 17:48.
 */
public class FormattedExporterTest {
	@Test
	public void test() throws Exception {
		Workbook workbook = new XSSFWorkbook();
		//model
		Model model = new Model();
		model.setName("学生统计");
		model.addAttribute(new Attribute("name", "学生名称", "String", "文本", null, null));
		model.addAttribute(new Attribute("number", "学生学号", "String", "文本", null, null));
		model.addAttribute(new Attribute("phone", "学生电话", "String", "文本", null, null));
		model.addAttribute(new Attribute("email", "学生电邮", "String", "文本", null, null));
		//data
		List<Model> datas = new ArrayList<>();
		for (int i = 0; i < 10; i++) {
			Model data = new Model();
			data.setName(model.getName());
			data.setAttributes(model.getAttributes());

			data.addAttrValue("学生" + i);
			data.addAttrValue("2017000" + i);
			data.addAttrValue("1500000000" + i);
			data.addAttrValue(i + "000@web.com");

			datas.add(data);
		}
		System.out
				.println(String.format("CreateResult : [%s]", FormattedExporter.createExcel(workbook, 0, model, "zh_CN")));
		System.out.println(String.format("InsertResult : [%s]", FormattedExporter.insertExcelData(workbook, 0, 6, datas)));

		//write file
		String filePath = "C:\\Users\\johnson\\Desktop\\test.xlsx";
		FileUtils.createFileDE(filePath);
		Common.writeFileToDisk(workbook, new File(filePath));
		Common.closeWorkbook(workbook);
	}
}
