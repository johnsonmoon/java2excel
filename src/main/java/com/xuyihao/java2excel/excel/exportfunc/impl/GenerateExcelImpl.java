package com.xuyihao.java2excel.excel.exportfunc.impl;

import com.xuyihao.java2excel.excel.exportfunc.ExportUtil;
import com.xuyihao.java2excel.model.ExcelTemplate;
import com.xuyihao.java2excel.excel.exportfunc.GenerateExcel;
import com.xuyihao.java2excel.model.ProgressMessage;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 */
public class GenerateExcelImpl implements GenerateExcel {

    @Autowired
    private ExportUtil exportUtil;

    public boolean generateTemplate(FileOutputStream fileOut, ExcelTemplate excelTemplate) {
        boolean flag = false;
        Workbook workbook = new HSSFWorkbook();
        int sheetNum = 0;
        ProgressMessage progressMessage = new ProgressMessage();
        progressMessage.reset();
        flag = this.exportUtil.createExcel(workbook, sheetNum, excelTemplate, true, fileOut, progressMessage);
        return flag;
    }

    public boolean generateFile(FileOutputStream fileOut, List<ExcelTemplate> dataList) {
        boolean flag = true;
        // 先写入模板

        //后写入具体数据

        return flag;
    }

    public ExportUtil getExportUtil() {
        return exportUtil;
    }

    public void setExportUtil(ExportUtil exportUtil) {
        this.exportUtil = exportUtil;
    }
}
