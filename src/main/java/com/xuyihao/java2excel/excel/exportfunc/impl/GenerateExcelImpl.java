package com.xuyihao.java2excel.excel.exportfunc.impl;

import com.xuyihao.java2excel.excel.exportfunc.util.ExportUtil;
import com.xuyihao.java2excel.excel.exportfunc.GenerateExcel;
import com.xuyihao.java2excel.excel.model.ExcelTemplate;
import com.xuyihao.java2excel.excel.model.ProgressMessage;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 *
 * 根据不同业务模型对本类进行重写
 */
public class GenerateExcelImpl implements GenerateExcel {

    @Autowired
    private ExportUtil exportUtil;

    public boolean generateTemplate(FileOutputStream fileOut, String classCode, ProgressMessage progressMessage) {
        return false;
    }

    public boolean generateTemplateWithMultiSheets(FileOutputStream fileOut, List<String> classCodeList, ProgressMessage progressMessage) {
        return false;
    }

    public boolean generateDataFile(FileOutputStream fileOut, String classCode, ProgressMessage progressMessage) {
        return false;
    }

    public boolean generateDataFileWithMutiSheets(FileOutputStream fileOut, List<String> classCodeList, ProgressMessage progressMessage) {
        return false;
    }

    public boolean generateSheet(Workbook workbook, int sheetNumber, FileOutputStream fileOut, String classCode, ProgressMessage progressMessage) {
        return false;
    }

    public ExcelTemplate getExcelTemplateByClassCode(String classCode) {
        return null;
    }

    public ExcelTemplate getExcelTemplateByModel(Object model) {
        return null;
    }

    public List<ExcelTemplate> convertModelToExcelTemplate(List<Object> modelList, ExcelTemplate template) {
        return null;
    }
}
