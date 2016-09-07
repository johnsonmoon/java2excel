package xuyihao.java2excel.excel;

import xuyihao.java2excel.excel.entity.ProgressMessage;
import xuyihao.java2excel.excel.util.CommonExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * Created by Xuyh at 2016/07/30 上午 11:09.
 *
 * 通过ProgressMessage的warningMessageList生成warning message文件，文件形式为excel表格
 */
public class WarningMessageFileGenerator {
    public boolean generateWarningFile(FileOutputStream fileOutputStream, ProgressMessage progressMessage){
        boolean flag;
        Workbook workbook = new XSSFWorkbook();
        int sheetNum = 0;
        flag = CommonExcelUtil.insertWarningMessageToSheet(workbook, sheetNum, true, fileOutputStream, progressMessage);
        return flag;
    }
}
