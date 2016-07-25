package com.xuyihao.java2excel.biz.excel.exportfunc.impl;

import com.xuyihao.java2excel.api.model.ExcelTemplate;
import com.xuyihao.java2excel.biz.excel.exportfunc.GenerateExcel;

import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 */
public class GenerateExcelImpl implements GenerateExcel {
    public boolean generateTemplate(FileOutputStream fileOut, ExcelTemplate excelTemplate) {
        return false;
    }

    public boolean generateFile(FileOutputStream fileOut, List<ExcelTemplate> dataList) {
        return false;
    }
}
