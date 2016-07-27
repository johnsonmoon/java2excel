package com.xuyihao.java2excel.excel.exportfunc.util;

import com.xuyihao.java2excel.excel.util.CommonExcelUtil;
import com.xuyihao.java2excel.excel.model.ExcelTemplate;
import com.xuyihao.java2excel.excel.model.ProgressMessage;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.util.*;

/**
 * Created by Xuyh at 2016/07/22 上午 11:36.
 *
 * Excel文件导出操作工具类
 */
public class ExportUtil extends CommonExcelUtil{
    /**
     * 创建表格
     *
     * @param workbook 表格
     * @param sheetNum 工作簿编号
     * @param excelTemplate 数据模型
     * @param ifCloseWorkBook 是否保存关闭表格
     * @param fileOut 文件流
     * @return true 成功, false 失败
     * */
    public boolean createExcel(final Workbook workbook, int sheetNum, ExcelTemplate excelTemplate, boolean ifCloseWorkBook, FileOutputStream fileOut, ProgressMessage progressMessage){
        boolean flag = false;
        progressMessage.reset();
        progressMessage.setDetailMessage("开始计算属性数量...");
        try{
            Sheet sheet = workbook.createSheet(excelTemplate.getClassName());
            workbook.setSheetOrder(excelTemplate.getClassName(), sheetNum);
            //总共有多少列数（后期添加关系等需要修改）
            int collumnSize = 0;
            int attrivalueSize = excelTemplate.getAttrbuteTypes().size();
            collumnSize = attrivalueSize + 1;
            //设置属性列宽
            if(excelTemplate != null){
                for(int i = 0; i < collumnSize; i++){
                    sheet.setColumnWidth(i+1, 3800);
                }
            }
            //隐藏列代码
            sheet.setColumnHidden(1, true);
            //隐藏行代码（设置行高为零）
            sheet.createRow(3).setZeroHeight(true);
            sheet.createRow(4).setZeroHeight(true);
            //合并单元格
            CellRangeAddress cellRangeAddress1 = new CellRangeAddress(0, 1, 0, 0);
            CellRangeAddress cellRangeAddress3 = new CellRangeAddress(0, 1, 1, collumnSize);
            sheet.addMergedRegion(cellRangeAddress1);
            sheet.addMergedRegion(cellRangeAddress3);
            //创建字体
            //Title
            Font fontTitle = workbook.createFont();
            fontTitle.setFontName(HSSFFont.FONT_ARIAL);
            fontTitle.setItalic(false);
            fontTitle.setFontHeightInPoints((short) 10);
            //Header
            Font fontHeader = workbook.createFont();
            fontHeader.setFontName(HSSFFont.FONT_ARIAL);
            fontHeader.setItalic(false);
            fontHeader.setFontHeightInPoints((short) 10);
            //HeaderFirstCell
            Font fontHeaderFirstCell = workbook.createFont();
            fontHeaderFirstCell.setFontName(HSSFFont.FONT_ARIAL);
            fontHeaderFirstCell.setColor(HSSFColor.WHITE.index);
            fontHeaderFirstCell.setItalic(false);
            fontHeaderFirstCell.setFontHeightInPoints((short) 10);
            //Value
            Font fontValue = workbook.createFont();
            fontValue.setFontName(HSSFFont.FONT_ARIAL);
            fontValue.setItalic(false);
            fontValue.setFontHeightInPoints((short) 10);

            // 创建单元格格式
            CellStyle cellStyleGeneral = workbook.createCellStyle();
            cellStyleGeneral.setWrapText(false);
            cellStyleGeneral.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

            CellStyle cellStyleTitle = workbook.createCellStyle();
            cellStyleTitle.cloneStyleFrom(cellStyleGeneral);
            cellStyleTitle.setFont(fontTitle);
            cellStyleTitle.setAlignment(CellStyle.ALIGN_CENTER);

            CellStyle cellStyleColumnHeader = workbook.createCellStyle();
            cellStyleColumnHeader.cloneStyleFrom(cellStyleGeneral);
            cellStyleColumnHeader.setFont(fontHeader);
            cellStyleColumnHeader.setAlignment(CellStyle.ALIGN_CENTER);

            CellStyle cellStyleRowHeader = workbook.createCellStyle();
            cellStyleRowHeader.cloneStyleFrom(cellStyleGeneral);
            cellStyleRowHeader.setFont(fontHeader);
            cellStyleRowHeader.setAlignment(CellStyle.ALIGN_CENTER);

            CellStyle cellStyleGrayRowHeader = workbook.createCellStyle();
            cellStyleGrayRowHeader.cloneStyleFrom(cellStyleRowHeader);

            CellStyle cellStyleRowHeaderTopAlign = workbook.createCellStyle();
            cellStyleRowHeaderTopAlign.cloneStyleFrom(cellStyleRowHeader);
            cellStyleRowHeaderTopAlign.setVerticalAlignment(CellStyle.VERTICAL_TOP);

            CellStyle cellStyleHideHeader = workbook.createCellStyle();
            cellStyleHideHeader.setFont(fontHeaderFirstCell);
            cellStyleHideHeader.setWrapText(true);
            cellStyleHideHeader.setAlignment(CellStyle.ALIGN_LEFT);
            cellStyleHideHeader.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            cellStyleHideHeader.setLocked(true);
            cellStyleHideHeader.setFillForegroundColor(IndexedColors.WHITE.getIndex());

            CellStyle cellStyleWhiteHideHeader = workbook.createCellStyle();
            cellStyleWhiteHideHeader.cloneStyleFrom(cellStyleHideHeader);
            cellStyleWhiteHideHeader.setFillForegroundColor(IndexedColors.WHITE.getIndex());

            CellStyle cellStyleValue = workbook.createCellStyle();
            cellStyleValue.cloneStyleFrom(cellStyleGeneral);
            cellStyleValue.setFont(fontValue);
            cellStyleValue.setAlignment(CellStyle.ALIGN_LEFT);
            cellStyleValue.setWrapText(true);
            cellStyleValue.setLocked(false);
            //写入固定数据
            String templateDetailInfo = excelTemplate.getId() + "&&" + excelTemplate.getTenant() + "&&" + excelTemplate.getClassCode() + "&&" + excelTemplate.getClassName();
            this.insertCellValue(sheet, 0, 0, templateDetailInfo, cellStyleHideHeader);
            this.insertCellValue(sheet, 1, 0, excelTemplate.getClassName(), cellStyleTitle);
            this.insertCellValue(sheet, 0, 2, "字段", cellStyleRowHeader);
            this.insertCellValue(sheet, 0, 3, "请勿编辑此行", cellStyleGrayRowHeader);
            this.insertCellValue(sheet, 0, 4, "数据格式", cellStyleGrayRowHeader);
            this.insertCellValue(sheet, 0, 5, "默认值", cellStyleRowHeader);
            this.insertCellValue(sheet, 0, 6, "数据", cellStyleRowHeaderTopAlign);
            // 增加資源标签列
            this.insertCellValue(sheet, 1, 2, "配置项标识", cellStyleColumnHeader);
            this.insertCellValue(sheet, 1, 3, "配置项标识", cellStyleWhiteHideHeader);
            //设置进度信息
            int totalCounts = excelTemplate.getAttrbuteTypes().size();
            progressMessage.setTotalCount(totalCounts);
            progressMessage.setDetailMessage("共有"+ totalCounts +"条数据");//设置进度为开始
            progressMessage.stateStarted();
            for(int i = 0; i < excelTemplate.getAttrbuteTypes().size(); i++){
                String label = excelTemplate.getAttrbuteTypes().get(i).getAttrName();
                String attrInfo = excelTemplate.getAttrbuteTypes().get(i).getAttrId() + "&&" + excelTemplate.getAttrbuteTypes().get(i).getAttrCode() + "&&" + excelTemplate.getAttrbuteTypes().get(i).getAttrName();
                String formatRule = excelTemplate.getAttrbuteTypes().get(i).getAttrFormatRule();
                String defaultValue = excelTemplate.getAttrbuteTypes().get(i).getDefaultValue();
                this.insertCellValue(sheet, i+2, 2, label, cellStyleGrayRowHeader);//字段名
                this.insertCellValue(sheet, i+2, 3, attrInfo, cellStyleGrayRowHeader);
                this.insertCellValue(sheet, i+2, 4, formatRule, cellStyleColumnHeader);//格式
                this.insertCellValue(sheet, i+2, 5, defaultValue, cellStyleColumnHeader);//默认值
                //设置进度信息
                progressMessage.addCurrentCount();
                progressMessage.setDetailMessage("进行  " + progressMessage.getCurrentCount() + "/" + progressMessage.getTotalCount() + "个操作");
            }
            flag = true;
        }catch(Exception e){
            e.printStackTrace();
            flag = false;
            progressMessage.stateFailed();
        }finally {
            if(workbook != null && ifCloseWorkBook){
                if(this.writeFileToDisk(workbook, fileOut)){
                    flag = true;
                }else {
                    progressMessage.stateFailed();
                }
            }
        }
        progressMessage.stateEnd();
        return flag;
    }

    /**
     *批量插入数据
     *
     * @param workbook 工作簿
     * @param sheetNum 表格编号
     * @param startRowNum 写入数据的起始行（第一次写入应当从第七行即startRowNum=6开始）
     * @param template 类型，用来定义表格属性（当excelTemplates数据列表为空时候无法通过列表检查类型属性）
     * @param excelTemplatesList 数据列表
     * @param ifCloseWorkBook 是否保存工作簿文件
     * @param fileOut 文件流
     * @return true 成功, false 失败
     */
    public boolean insertExcelData(final Workbook workbook, int sheetNum, int startRowNum, ExcelTemplate template, List<ExcelTemplate> excelTemplatesList, boolean ifCloseWorkBook, FileOutputStream fileOut, ProgressMessage progressMessage) {
        boolean flag = false;
        progressMessage.reset();
        progressMessage.setDetailMessage("开始计算数据数量...");
        if (workbook == null) {
            return false;
        }
        Sheet sheet = workbook.getSheetAt(sheetNum);
        if (sheet == null) {
            return false;
        }
        String[] identifyString = this.getCellValue(sheet, 0, 0).split("&&");
        if (!identifyString[2].equals(template.getClassCode()) || !identifyString[0].equals(template.getId())) {
            return false;//excel对应的类型和传入的excelTemplates类型不一致
        }
        //设置字体
        Font fontHeader = workbook.createFont();
        fontHeader.setFontName(HSSFFont.FONT_ARIAL);
        fontHeader.setItalic(false);
        fontHeader.setFontHeightInPoints((short) 10);
        Font fontValue = fontHeader;
        try {
            // 创建单元格格式
            CellStyle cellStyleValue = workbook.createCellStyle();
            cellStyleValue.setWrapText(true);
            cellStyleValue.setFont(fontValue);
            cellStyleValue.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            cellStyleValue.setAlignment(CellStyle.ALIGN_LEFT);
            cellStyleValue.setLocked(false);
            //设置进度信息
            int totalCounts = excelTemplatesList.size();//CI的数量
            progressMessage.setTotalCount(totalCounts);
            progressMessage.setDetailMessage("共有" + totalCounts + "条数据");
            //设置进度为开始
            progressMessage.stateStarted();
            if (startRowNum < 6) {
                System.out.println("至少从第7(6+1)行开始插入数据!");
                progressMessage.stateFailed();
            }
            for (int j = 0; j < excelTemplatesList.size(); j++) {
                for (int i = 0; i < excelTemplatesList.get(j).getAttrValues().size(); i++) {
                    this.insertCellValue(sheet, i + 2, j + startRowNum, excelTemplatesList.get(j).getAttrValues().get(i), cellStyleValue);
                }
                //设置进度信息
                progressMessage.addCurrentCount();
                progressMessage.setDetailMessage("进行  " + progressMessage.getCurrentCount() + "/" + progressMessage.getTotalCount() + "个操作");
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("写入excel表格数据失败");
            flag = false;
            progressMessage.stateFailed();
        } finally {
            if (workbook != null && ifCloseWorkBook) {
                if (this.writeFileToDisk(workbook, fileOut)) {
                    flag = true;
                } else {
                    progressMessage.stateFailed();
                }
            }
        }
        progressMessage.stateEnd();
        return flag;
    }
}
