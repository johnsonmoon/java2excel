package com.xuyihao.java2excel.biz.excel.exportfunc;

import com.xuyihao.java2excel.api.model.ExcelTemplate;

import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 *
 * 生成Excel表格接口
 */
public interface GenerateExcel {
    /**
     *通过ExcelTemplate生成表格模板
     *
     * @param fileOut
     * @param excelTemplate
     * @return false 失败， true 成功
     * */
    public boolean generateTemplate(FileOutputStream fileOut, ExcelTemplate excelTemplate);

    /**
     *连同数据一起生成完整表格
     *
     * @param fileOut
     * @param dataList
     * @return false 失败， true 成功
     * */
    public boolean generateFile(FileOutputStream fileOut, List<ExcelTemplate> dataList);
}
