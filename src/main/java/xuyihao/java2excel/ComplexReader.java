package xuyihao.java2excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * ComplexReader
 * TODO
 * <p>
 * <p>
 * <pre>
 * 	Read excel file at multiple sheets.
 * </pre>
 * <p>
 * Created by xuyh at 2018/1/5 16:52.
 */
public class ComplexReader extends AbstractReader {
	private static Logger logger = LoggerFactory.getLogger(ComplexReader.class);
	private Workbook workbook;

	public ComplexReader() {
	}

	public ComplexReader(String filePathName) {
		openFile(filePathName);
	}

	public boolean openFile(String filePathName) {
		try {
			if (workbook != null) {
				close(workbook);
			}
			workbook = new XSSFWorkbook(filePathName);
			return true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
	}

	public boolean close() {
		if (workbook == null)
			return false;
		return close(workbook);
	}
}
