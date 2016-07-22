package com.xuyihao.java2excel.biz.excel.exportfunc;

import com.xuyihao.java2excel.api.model.AttributeType;
import com.xuyihao.java2excel.api.model.ExcelTemplate;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.*;

/**
 * Created by Xuyh at 2016/07/22 上午 11:36.
 *
 * Excel文件操作工具类
 */
public class exportUtil {
    /**
     * 创建表格
     *
     * @param workbook 表格
     * @param sheetNum 工作簿编号
     * @param excelTemplate 数据模型
     * @param ifCloseWorkBook 是否保存关闭表格
     * @param fileOut 文件流
     *@return true 成功, false 失败
     * */
    public boolean createExcel(final Workbook workbook, int sheetNum, ExcelTemplate excelTemplate, boolean ifCloseWorkBook, FileOutputStream fileOut){
        boolean flag = false;
        try{
            Sheet sheet = workbook.createSheet(excelTemplate.getClassName());
            workbook.setSheetOrder(excelTemplate.getClassName(), sheetNum);
            //总共有多少列数（后期添加关系等需要修改）
            int collumnSize = 0;
            int attrivalueSize = excelTemplate.getAttrValues().size();
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
            //合并单元格
            List<MergeRange> mergeRanges = new ArrayList<MergeRange>();
            mergeRanges.add(new MergeRange(0, 1, 0, 0));
            mergeRanges.add(new MergeRange(0, 1, 1, collumnSize + 1));
            this.mergeCellsByMulti(sheet, mergeRanges);
            //创建字体
            //Title
            Font fontTitle = workbook.createFont();
            fontTitle.setFontName(HSSFFont.FONT_ARIAL);
            fontTitle.setItalic(false);
            //fontTitle.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
            fontTitle.setFontHeightInPoints((short) 10);

            //Header
            Font fontHeader = workbook.createFont();
            fontHeader.setFontName(HSSFFont.FONT_ARIAL);
            fontHeader.setItalic(false);
            //fontHeader.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
            fontHeader.setFontHeightInPoints((short) 10);

            //HeaderFirstCell
            Font fontHeaderFirstCell = workbook.createFont();
            fontHeaderFirstCell.setFontName(HSSFFont.FONT_ARIAL);
            fontHeaderFirstCell.setItalic(false);
            //fontHeaderFirstCell.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);
            fontHeaderFirstCell.setFontHeightInPoints((short) 10);
            //fontHeaderFirstCell.setColor(IndexedColors.WHITE.getIndex());

            Font fontValue = fontHeader;

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
            this.insertCellValue(sheet, 0, 0, excelTemplate.getId() + "&&" +  excelTemplate.getTenant() + "&&" + excelTemplate.getClassCode() + "&&" + excelTemplate, cellStyleHideHeader);
            this.insertCellValue(sheet, 1, 0, "配置项模板表格（" + excelTemplate.getClassName() + "）", cellStyleTitle);
            this.insertCellValue(sheet, 0, 2, "字段", cellStyleRowHeader);
            this.insertCellValue(sheet, 0, 3, "请勿编辑此行", cellStyleGrayRowHeader);
            this.insertCellValue(sheet, 0, 4, "数据格式", cellStyleRowHeader);
            this.insertCellValue(sheet, 0, 5, "默认值", cellStyleRowHeader);
            this.insertCellValue(sheet, 0, 6, "配置项数据", cellStyleRowHeaderTopAlign);
            // 增加資源标签列
            this.insertCellValue(sheet, 1, 2, "配置项标识", cellStyleColumnHeader);
            this.insertCellValue(sheet, 1, 3, "配置项标识", cellStyleWhiteHideHeader);
            int i = 2;
            Iterator it = excelTemplate.getAttrValues().entrySet().iterator();
            while(it.hasNext()){
                Map.Entry entity = (Map.Entry)it.next();
                AttributeType attributeType = (AttributeType)entity.getKey();
                String label = attributeType.getAttrName();
                this.insertCellValue(sheet, );


                i++;
            }
        }catch(IOException e){
            e.printStackTrace();
            flag = false;
        }finally {
            if(workbook != null && ifCloseWorkBook){
                try{
                    workbook.write(fileOut);
                }catch (Exception e1){
                    e1.printStackTrace();
                    flag = false;
                }
            }
        }
        return flag;
    }

    /**
     * 插入单元格数据
     *
     * @param sheet 工作簿
     * @param collumn 列
     * @param row 行
     * @param value 值
     * @param cellStyle 单元格样式风格
     * */
    private void insertCellValue(Sheet sheet, int collumn, int row, String value, CellStyle cellStyle){
        Row targetRow = sheet.getRow(row);
        if(targetRow == null){
            targetRow = sheet.createRow(row);
        }
        if(targetRow != null){
            Cell cell = targetRow.createCell(collumn);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(value);
        }
    }

    /**
     * 获取单元格值
     *
     * @param sheet 工作簿
     * @param row 行
     * @param collumn 列
     * @return 单元格值
     * */
    private String getCellValue(Sheet sheet, int row, int collumn){
        String cellValue = "";
        Row targetRow = sheet.getRow(row);
        if(targetRow == null){
            cellValue = "";
        }else{
            Cell cell = targetRow.getCell(collumn);
            if(cell == null){
                cellValue = "";
            }else{
                DecimalFormat decimalFormat = new DecimalFormat("#");
                switch (cell.getCellType()){
                    case Cell.CELL_TYPE_STRING:
                        cellValue = cell.getRichStringCellValue().getString().trim();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        cellValue = decimalFormat.format(cell.getNumericCellValue()).toString();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        cellValue = cell.getCellFormula();
                        break;
                    default:
                        cellValue = "";
                }
            }
        }
        return cellValue;
    }

    /**
     * 合并单元格
     *
     *@param mergeRanges 需要合并的区域集合
     *@param sheet 需要操作的工作薄
     * */
    public void mergeCellsByMulti(Sheet sheet, List<MergeRange> mergeRanges){
        for(MergeRange range : mergeRanges){
            CellRangeAddress region = new CellRangeAddress(range.firstRow, range.lastRow, range.firstCollumn, range.lastCollum);
            sheet.addMergedRegion(region);
        }
    }

    /**
     * inner class
     *
     * 定位合并单元格的范围（区域）
     * */
    public class MergeRange{
        public int firstRow;
        public int lastRow;
        public int firstCollumn;
        public int lastCollum;

        public mergeRange(int firstRow, int lastRow, int firstCollumn, int lastCollum){
            this.firstCollumn = firstRow;
            this.lastCollum = lastCollum;
            this.firstRow = firstRow;
            this.lastRow = lastRow;
        }
    }
}
