package com.xuyihao.java2excel.excel.exportfunc;

import com.xuyihao.java2excel.excel.model.ProgressMessage;

import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 *
 * 生成Excel表格接口
 * 需要对不同的项目业务模型进行适配
 * 不同项目可以根据不同需求对实现进行重写
 */
public interface GenerateExcel {
    /**
     * 通过ExcelTemplate生成表格模板
     *
     * @param fileOut 文件输出流
     * @param classCode 类编码
     * @param progressMessage 进度通知
     * @return false 失败， true 成功
     * */
    public boolean generateTemplate(FileOutputStream fileOut, String classCode, ProgressMessage progressMessage);

    /**
     * 连同数据一起生成完整表格,一个类生成一个excel文件
     *
     * @param fileOut 文件输出流
     * @param classCode 类编码
     * @param progressMessage 进度通知
     * @return false 失败， true 成功
     * */
    public boolean generateFile(FileOutputStream fileOut, String classCode, ProgressMessage progressMessage);

    /**
     * 连同数据一起生成完整表格,并将类列表的所有类数据全部放在同一张excel的不同sheet中
     *
     * @param fileOut 文件输出流
     * @param classCodeList 类编码列表
     * @param progressMessage 进度通知
     * @return false 失败， true 成功
     */
    public boolean generateFileWithMutiSheets(FileOutputStream fileOut, List<String> classCodeList, ProgressMessage progressMessage);
}
