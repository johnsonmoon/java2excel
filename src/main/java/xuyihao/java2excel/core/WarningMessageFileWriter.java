package xuyihao.java2excel.core;

import xuyihao.java2excel.core.model.ProgressMessage;
import xuyihao.java2excel.core.util.CommonExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * Created by Xuyh at 2016/07/30 上午 11:09.
 *
 * 通过ProgressMessage的warningMessageList生成warning message文件，文件形式为excel表格
 */
public class WarningMessageFileWriter {
    public boolean generateWarningFile(FileOutputStream fileOutputStream, ProgressMessage progressMessage){
        boolean flag;
        Workbook workbook = new XSSFWorkbook();
        int sheetNum = 0;
        flag = CommonExcelUtil.insertWarningMessageToSheet(workbook, sheetNum, true, fileOutputStream, progressMessage);
        return flag;
    }
}
