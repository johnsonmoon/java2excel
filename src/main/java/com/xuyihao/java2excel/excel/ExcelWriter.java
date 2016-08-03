package com.xuyihao.java2excel.excel;

import com.xuyihao.java2excel.excel.model.ExcelTemplate;
import com.xuyihao.java2excel.excel.model.ProgressMessage;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 *
 * 生成Excel表格接口
 * 需要对不同的项目业务模型进行适配
 * 不同项目可以根据不同需求对实现进行重写
 *
 * 重写逻辑：
 * 1.generateTemplate，generateTemplateWithMultiSheets方法，设置状态信息（包括detailMessage, state, totalCount），新建工作簿对象，指定sheet编号，
 *   重写getExcelTemplate两个方法，并获取ExcelTemplate对象，调用ExportUtil的createExcel方法，过程中设置状态信息的currentCount，结束后设置state为end。
 * 2.generateFile，generateFileWithMutiSheets方法，设置状态信息，新建工作簿，重写generateSheet方法并调用之。
 * 3.generateSheet方法逻辑：分析模型，写入模板（调用ExportUtil的createExcel方法），
 *   写入数据（调用convertModelToExcelTemplate方法获取ExcelTemplate对象列表，调用ExportUtil的insertExcelData插入数据）
 */
public interface ExcelWriter {
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
     * 通过classCode生成表格模板,将类列表的所有类数据全部放在同一张excel的不同sheet中
     *
     * @param fileOut 文件输出流
     * @param classCodeList 类编码列表
     * @param progressMessage 进度通知
     * @return false 失败， true 成功
     */
    public boolean generateTemplateWithMultiSheets(FileOutputStream fileOut, List<String> classCodeList, ProgressMessage progressMessage);

    /**
     * 连同数据一起生成完整表格,一个类生成一个excel文件
     *
     * @param fileOut 文件输出流
     * @param classCode 类编码
     * @param progressMessage 进度通知
     * @return false 失败， true 成功
     * */
    public boolean generateDataFile(FileOutputStream fileOut, String classCode, ProgressMessage progressMessage);

    /**
     * 连同数据一起生成完整表格,并将类列表的所有类数据全部放在同一张excel的不同sheet中
     *
     * @param fileOut 文件输出流
     * @param classCodeList 类编码列表
     * @param progressMessage 进度通知
     * @return false 失败， true 成功
     */
    public boolean generateDataFileWithMutiSheets(FileOutputStream fileOut, List<String> classCodeList, ProgressMessage progressMessage);

    /**
     * 生成单个sheet
     *
     * @param workbook excel表格
     * @param sheetNumber sheet编号
     * @param fileOut 文件输出流
     * @param classCode ExcelTemplate类编号
     * @param progressMessage 消息进度编号
     * @return false 失败， true 成功
     */
    public boolean generateSheet(Workbook workbook, int sheetNumber, FileOutputStream fileOut, String classCode, ProgressMessage progressMessage);

    /**
     * 通过类编码查询数据库并对对应的ExcelTemplate进行赋值,需要根据不同的业务模型进行重写
     *
     * @param classCode 模型类型编码
     * @return ExcelTemplate对象
     */
    public ExcelTemplate getExcelTemplateByClassCode(String classCode);

    /**
     * 通过具体的业务模型对象对ExcelTemplate进行赋值，这里需要进行重写
     *
     * @param model 业务模型对象
     * @return ExcelTemplate对象
     */
    public ExcelTemplate getExcelTemplateByModel(Object model);

    /**
     * 将模型对象列表转化为ExcelTemplate对象列表，需要对不同业务模型进行重写
     *
     * @param modelList 业务模型对象列表
     * @param template 已经转化过的单个ExcelTemplate对象
     * @return ExcelTemplate对象列表
     */
    public List<ExcelTemplate> convertModelToExcelTemplate(List<Object> modelList, ExcelTemplate template);
}
