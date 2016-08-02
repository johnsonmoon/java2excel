package com.xuyihao.java2excel.excel.importfunc.impl;

import com.xuyihao.java2excel.excel.importfunc.ReadExcel;
import com.xuyihao.java2excel.excel.model.ExcelTemplate;
import com.xuyihao.java2excel.excel.model.ProgressMessage;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/27 下午 02:56.
 */
public class ReadExcelImpl implements ReadExcel{
    public boolean processDatas(FileInputStream fileIn, ProgressMessage progressMessage) {
        return false;
    }

    public void processSingleSheetData(Workbook workbook, int sheetNumber, ProgressMessage progressMessage) {

    }

    public void processExcelTemplateList(List<ExcelTemplate> excelTemplateList, int sheetNumber, ProgressMessage progressMessage, int currentStartRow) {

    }
}
