package com.xuyihao.java2excel.excel.util;

import com.xuyihao.java2excel.excel.entity.ProgressMessage;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;

/**
 * Created by Xuyh at 2016/07/27 下午 02:28.
 */
public class CommonExcelUtil {
    /**
     * 插入单元格数据
     *
     * @param sheet 工作簿
     * @param collumn 列
     * @param row 行
     * @param value 值
     * @param cellStyle 单元格样式风格
     * */
    public static void insertCellValue(Sheet sheet, int collumn, int row, String value, CellStyle cellStyle){
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
    public static String getCellValue(Sheet sheet, int row, int collumn){
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
     * 将workbook写入磁盘
     *
     * @param workbook excel表格
     * @param fileOut 文件输出流
     * @return true if succeeded, false if failed
     */
    public static boolean writeFileToDisk(Workbook workbook, FileOutputStream fileOut){
        boolean flag = false;
        try{
            workbook.write(fileOut);
            flag = true;
        }catch (IOException e1){
            e1.printStackTrace();
            flag = false;
        }
        return flag;
    }

    /**
     * 编写警告信息汇总表格并保存到磁盘
     *
     * @param workbook 表格
     * @param sheetNumber 工作簿编号
     * @param ifCloseWorkBook 是否将文件保存到磁盘
     * @param fileOut 文件输出流
     * @param progressMessage 进度消息
     * @return
     */
    public static boolean insertWarningMessageToSheet(final Workbook workbook, int sheetNumber, boolean ifCloseWorkBook, FileOutputStream fileOut, ProgressMessage progressMessage){
        boolean flag = true;
        if(workbook == null){
            return false;
        }
        Sheet sheet = workbook.createSheet("Warning Message Summary");
        //设置字体
        Font fontHeader = workbook.createFont();
        fontHeader.setFontName(HSSFFont.FONT_ARIAL);
        fontHeader.setItalic(false);
        fontHeader.setFontHeightInPoints((short) 10);
        Font fontValue = fontHeader;
        try{
            progressMessage.setDetailMessage("Generating warning message file...");
            // 创建单元格格式
            CellStyle cellStyleValue = workbook.createCellStyle();
            cellStyleValue.setWrapText(true);
            cellStyleValue.setFont(fontValue);
            cellStyleValue.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            cellStyleValue.setAlignment(CellStyle.ALIGN_LEFT);
            cellStyleValue.setLocked(false);
            //设置列宽
            sheet.setColumnWidth(0, 40000);
            //合并单元格
            CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 1);
            sheet.addMergedRegion(cellRangeAddress);
            insertCellValue(sheet, 0, 0, "警告信息汇总", cellStyleValue);
            insertCellValue(sheet, 0, 1, "信息", cellStyleValue);
            //插入数据
            for (int j = 0; j < progressMessage.getWarningMessageList().size(); j++){
                insertCellValue(sheet, 0, j+2, progressMessage.getWarningMessageList().get(j), cellStyleValue);
            }
        }catch (Exception e){
            e.printStackTrace();
            progressMessage.setDetailMessage("Failed generating file!");
            flag = false;
            progressMessage.stateFailed();
        }
        finally {
            progressMessage.setDetailMessage("Succeeded generating file!");
            if(workbook != null && ifCloseWorkBook){
                if(writeFileToDisk(workbook, fileOut)){
                    flag = true;
                }else {
                    progressMessage.stateFailed();
                }
            }
        }
        return flag;
    }
}
