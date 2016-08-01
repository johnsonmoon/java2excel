package com.xuyihao.java2excel.excel.importfunc;

import com.sun.javaws.progress.Progress;
import com.xuyihao.java2excel.excel.model.ProgressMessage;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;

/**
 * Created by Xuyh at 2016/07/27 下午 02:56.
 *
 * 读取Excel表格接口
 * 根据不同系统模型适配实现该接口
 */
public interface ReadExcel {
    /**
     * 读取excel文件内容，并且将数据保存到数据库
     * 根据不同业务系统进行重写
     *
     * @param fileIn 文件输入流
     * @param progressMessage 进度消息
     * @return false 失败， true 成功
     */
    public boolean saveDatas(FileInputStream fileIn, ProgressMessage progressMessage);

    /**
     * 对excel表中的单个sheet的数据进行处理
     *
     * @param workbook excel表格
     * @param sheetNumber sheet编号
     * @param progressMessage 进度消息
     */
    public void processSingleSheetData(Workbook workbook, int sheetNumber, ProgressMessage progressMessage);
}
