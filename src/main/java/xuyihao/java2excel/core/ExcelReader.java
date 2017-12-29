package xuyihao.java2excel.core;

import xuyihao.java2excel.core.entity.model.Template;
import xuyihao.java2excel.core.entity.ProgressMessage;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/27 下午 02:56.
 * <p>
 * 读取Excel表格接口
 * 根据不同系统模型适配实现该接口
 * <pre>
 * 重写逻辑：
 * 1.processExcelTemplateList方法：逐个解析ExcelTemplate对象列表为实际业务模型对象，并将其保存至系统数据库，并设置进度信息currentCount++
 *   过程中出错，则调用ProgressMessage的addWarningMessage方法添加警告信息。
 * 2.processSingleSheetData方法：逐个解析workbook中的每个sheet，通过ImportUtil的getExcelTemplateListDataFromExcel方法获取ExcelTemplate对象列表，
 *   并调用processExcelTemplateList方法进行处理（过程中要设置当前sheet读取的行，用来定位错误消息的具体sheet编号和行）。
 * 3.processDatas方法：设置进度信息，通过FileInputStream获取Workbook，通过ImportUtil的getExcelTemplateFromExcel方法获取基本的ExcelTEmplate对象信息，
 *   获取sheet数量，设置progressMessage的totalCount（对逐个sheet调用ImportUtil的getAttrValueCount获取单个sheet数据量）
 *   最后循环调用processSingleSheetData方法处理数据
 *   </pre>
 */
public interface ExcelReader {
	/**
	 * 读取excel文件内容，并且将数据保存到数据库
	 * 根据不同业务系统进行重写
	 *
	 * @param fileIn          文件输入流
	 * @param progressMessage 进度消息
	 * @return false 失败， true 成功
	 */
	public boolean processDatas(FileInputStream fileIn, ProgressMessage progressMessage);

	/**
	 * 对excel表中的单个sheet的数据进行处理
	 *
	 * @param workbook        excel表格
	 * @param sheetNumber     sheet编号
	 * @param progressMessage 进度消息
	 */
	public void processSingleSheetData(Workbook workbook, int sheetNumber, ProgressMessage progressMessage);

	/**
	 * 逐个处理ExcelTemplateList中的数据
	 *
	 * @param templateList    ExcelTemplate对象列表
	 * @param sheetNumber     工作簿编号
	 * @param progressMessage 进度消息对象
	 * @param currentStartRow 在excel文件的当前sheet中的当前开始的行，用来给ProgressMessage对象进行currentCount赋值进行计数
	 */
	public void processExcelTemplateList(List<Template> templateList, int sheetNumber, ProgressMessage progressMessage,
			int currentStartRow);
}
