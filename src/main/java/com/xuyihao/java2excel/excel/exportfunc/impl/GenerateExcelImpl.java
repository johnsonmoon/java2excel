package com.xuyihao.java2excel.excel.exportfunc.impl;

import com.xuyihao.java2excel.excel.exportfunc.util.ExportUtil;
import com.xuyihao.java2excel.excel.exportfunc.GenerateExcel;
import com.xuyihao.java2excel.excel.model.ProgressMessage;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 */
public class GenerateExcelImpl implements GenerateExcel {

    @Autowired
    private ExportUtil exportUtil;

    public boolean generateTemplate(FileOutputStream fileOut, String classCode, ProgressMessage progressMessage) {
        return false;
    }

    public boolean generateFile(FileOutputStream fileOut, String classCode, ProgressMessage progressMessage) {
        return false;
    }

    public boolean generateFileWithMutiSheets(FileOutputStream fileOut, List<String> classCodeList, ProgressMessage progressMessage) {
        return false;
    }

    public ExportUtil getExportUtil() {
        return exportUtil;
    }

    public void setExportUtil(ExportUtil exportUtil) {
        this.exportUtil = exportUtil;
    }
}
