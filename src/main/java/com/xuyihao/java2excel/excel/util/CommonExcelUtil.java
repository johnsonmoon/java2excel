package com.xuyihao.java2excel.excel.util;

import org.apache.poi.ss.usermodel.*;

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
    protected void insertCellValue(Sheet sheet, int collumn, int row, String value, CellStyle cellStyle){
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
    protected String getCellValue(Sheet sheet, int row, int collumn){
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
    protected boolean writeFileToDisk(Workbook workbook, FileOutputStream fileOut){
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
}
