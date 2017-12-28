package xuyihao.java2excel.core.operation;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.text.DecimalFormat;

/**
 * Created by Xuyh at 2016/07/27 下午 02:28.
 */
public class Common {
    private static Logger logger = LoggerFactory.getLogger(Common.class);

    public static final int CELL_STYLE_TYPE_GENERAL = 0;
    public static final int CELL_STYLE_TYPE_TITLE = 1;
    public static final int CELL_STYLE_TYPE_COLUMN_HEADER = 2;
    public static final int CELL_STYLE_TYPE_ROW_HEADER = 3;
    public static final int CELL_STYLE_TYPE_ROW_HEADER_GRAY = 4;
    public static final int CELL_STYLE_TYPE_ROW_HEADER_TOP_ALIGN = 5;
    public static final int CELL_STYLE_TYPE_HEADER_HIDE = 6;
    public static final int CELL_STYLE_TYPE_HEADER_HIDE_WHITE = 7;
    public static final int CELL_STYLE_TYPE_VALUE = 8;

    /**
     * 创建ARIAL字体
     *
     * @param workbook excel表格
     * @param size     字体大小
     * @return
     */
    public static Font createFontArial(Workbook workbook, int size) {
        Font font = workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setItalic(false);
        font.setFontHeightInPoints((short) size);
        return font;
    }

    /**
     * 创建ARIAL字体
     *
     * @param workbook excel表格
     * @param size     字体大小
     * @return
     */
    public static Font createFontWhiteArial(Workbook workbook, int size) {
        Font font = workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setColor(HSSFColor.WHITE.index);
        font.setItalic(false);
        font.setFontHeightInPoints((short) size);
        return font;
    }

    /**
     * 创建单元格格式
     *
     * @param workbook      excel表格
     * @param cellStyleType 格式类型
     * @return
     */
    public static CellStyle createCellStyle(Workbook workbook, int cellStyleType) {
        CellStyle cellStyleGeneral = workbook.createCellStyle();
        cellStyleGeneral.setWrapText(false);
        cellStyleGeneral.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        CellStyle cellStyleTitle = workbook.createCellStyle();
        cellStyleTitle.cloneStyleFrom(cellStyleGeneral);
        cellStyleTitle.setFont(createFontArial(workbook, 10));
        cellStyleTitle.setAlignment(CellStyle.ALIGN_CENTER);

        CellStyle cellStyleColumnHeader = workbook.createCellStyle();
        cellStyleColumnHeader.cloneStyleFrom(cellStyleGeneral);
        cellStyleColumnHeader.setFont(createFontArial(workbook, 10));
        cellStyleColumnHeader.setAlignment(CellStyle.ALIGN_CENTER);

        CellStyle cellStyleRowHeader = workbook.createCellStyle();
        cellStyleRowHeader.cloneStyleFrom(cellStyleGeneral);
        cellStyleRowHeader.setFont(createFontArial(workbook, 10));
        cellStyleRowHeader.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleRowHeader.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        cellStyleRowHeader.setFillPattern(CellStyle.SOLID_FOREGROUND);

        CellStyle cellStyleGrayRowHeader = workbook.createCellStyle();
        cellStyleGrayRowHeader.cloneStyleFrom(cellStyleRowHeader);

        CellStyle cellStyleRowHeaderTopAlign = workbook.createCellStyle();
        cellStyleRowHeaderTopAlign.cloneStyleFrom(cellStyleRowHeader);
        cellStyleRowHeaderTopAlign.setVerticalAlignment(CellStyle.VERTICAL_TOP);

        CellStyle cellStyleHideHeader = workbook.createCellStyle();
        cellStyleHideHeader.setFont(createFontWhiteArial(workbook, 10));
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
        cellStyleValue.setFont(createFontArial(workbook, 10));
        cellStyleValue.setAlignment(CellStyle.ALIGN_LEFT);
        cellStyleValue.setWrapText(true);
        cellStyleValue.setLocked(false);

        switch (cellStyleType) {
            case CELL_STYLE_TYPE_GENERAL:
                return cellStyleGeneral;
            case CELL_STYLE_TYPE_TITLE:
                return cellStyleTitle;
            case CELL_STYLE_TYPE_COLUMN_HEADER:
                return cellStyleColumnHeader;
            case CELL_STYLE_TYPE_ROW_HEADER:
                return cellStyleRowHeader;
            case CELL_STYLE_TYPE_ROW_HEADER_GRAY:
                return cellStyleGrayRowHeader;
            case CELL_STYLE_TYPE_ROW_HEADER_TOP_ALIGN:
                return cellStyleRowHeaderTopAlign;
            case CELL_STYLE_TYPE_HEADER_HIDE:
                return cellStyleHideHeader;
            case CELL_STYLE_TYPE_HEADER_HIDE_WHITE:
                return cellStyleWhiteHideHeader;
            case CELL_STYLE_TYPE_VALUE:
                return cellStyleValue;
            default:
                return cellStyleValue;
        }
    }


    /**
     * 插入单元格数据
     *
     * @param sheet     工作表
     * @param column    列
     * @param row       行
     * @param value     值
     * @param cellStyle 单元格样式风格
     */
    public static void insertCellValue(Sheet sheet, int column, int row, String value, CellStyle cellStyle) {
        Row targetRow = sheet.getRow(row);
        if (targetRow == null) {
            targetRow = sheet.createRow(row);
        }
        if (targetRow != null) {
            Cell cell = targetRow.createCell(column);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(value);
        }
    }

    /**
     * 获取单元格值
     *
     * @param sheet  工作表
     * @param row    行
     * @param column 列
     * @return 单元格值
     */
    public static String getCellValue(Sheet sheet, int row, int column) {
        String cellValue;
        Row targetRow = sheet.getRow(row);
        if (targetRow == null) {
            cellValue = "";
        } else {
            Cell cell = targetRow.getCell(column);
            if (cell == null) {
                cellValue = "";
            } else {
                DecimalFormat decimalFormat = new DecimalFormat("#");
                switch (cell.getCellType()) {
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
     * @param fileOut  文件输出流
     * @return true if succeeded, false if failed
     * @throws Exception
     */
    public static boolean writeFileToDisk(Workbook workbook, FileOutputStream fileOut) {
        try {
            workbook.write(fileOut);
        } catch (Exception e) {
            logger.warn(e.getMessage(), e);
            return false;
        }
        return true;
    }
}
