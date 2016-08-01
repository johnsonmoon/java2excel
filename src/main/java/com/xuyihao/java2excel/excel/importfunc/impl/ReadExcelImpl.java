package com.xuyihao.java2excel.excel.importfunc.impl;

import com.xuyihao.java2excel.excel.importfunc.ReadExcel;
import com.xuyihao.java2excel.excel.model.ProgressMessage;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;

/**
 * Created by Xuyh at 2016/07/27 下午 02:56.
 */
public class ReadExcelImpl implements ReadExcel{
    public boolean saveDatas(FileInputStream fileIn, ProgressMessage progressMessage) {
        boolean flag = false;
        return flag;
    }

    public void processSingleSheetData(Workbook workbook, int sheetNumber, ProgressMessage progressMessage) {

    }
}
