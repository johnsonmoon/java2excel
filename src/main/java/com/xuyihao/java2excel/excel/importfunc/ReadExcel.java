package com.xuyihao.java2excel.excel.importfunc;

import com.xuyihao.java2excel.excel.model.ProgressMessage;

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
     *
     * @param fileIn 文件输入流
     * @param progressMessage 进度消息
     * @return
     */
    public boolean saveDatas(FileInputStream fileIn, ProgressMessage progressMessage);
}
