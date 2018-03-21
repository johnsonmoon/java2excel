package xuyihao.java2excel.core.operation;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import xuyihao.java2excel.core.entity.common.CellStyle;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Xuyh at 2016/07/27 下午 02:28.
 */
public class Common {
	private static Logger logger = LoggerFactory.getLogger(Common.class);

	/**
	 * cell style cache for caching cell styles
	 * <pre>
	 *     We can define only up to 64000 style in a .xlsx Workbook.
	 *     So cell styles for workbook should be cached.
	 *     After workbook has been closed, cache needed to be cleaned.
	 * </pre>
	 */
	private static Map<Workbook, Map<Integer, org.apache.poi.ss.usermodel.CellStyle>> cellStyleCache = new HashMap<>();

	/**
	 * create ARIAL font
	 *
	 * @param workbook excel workbook
	 * @param size     font size
	 * @return Font
	 */
	public static Font createFontArial(Workbook workbook, int size) {
		Font font = workbook.createFont();
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setItalic(false);
		font.setFontHeightInPoints((short) size);
		return font;
	}

	/**
	 * create ARIAL font with white color
	 *
	 * @param workbook excel workbook
	 * @param size     font size
	 * @return Font
	 */
	public static Font createFontWhiteArial(Workbook workbook, int size) {
		Font font = workbook.createFont();
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
		font.setItalic(false);
		font.setFontHeightInPoints((short) size);
		return font;
	}

	/**
	 * create cell style
	 *
	 * @param workbook      excel workbook
	 * @param cellStyleType {Common#CELL_STYLE_TYPE_*}
	 * @return cell style
	 */
	public static org.apache.poi.ss.usermodel.CellStyle createCellStyle(Workbook workbook, int cellStyleType) {
		if (!cellStyleCache.containsKey(workbook)) {
			Map<Integer, org.apache.poi.ss.usermodel.CellStyle> styleMap = new HashMap<>();

			org.apache.poi.ss.usermodel.CellStyle cellStyleGeneral = workbook.createCellStyle();
			cellStyleGeneral.setWrapText(false);
			cellStyleGeneral.setVerticalAlignment(VerticalAlignment.CENTER);
			styleMap.put(CellStyle.CELL_STYLE_TYPE_GENERAL.getCode(), cellStyleGeneral);

			org.apache.poi.ss.usermodel.CellStyle cellStyleTitle = workbook.createCellStyle();
			cellStyleTitle.cloneStyleFrom(cellStyleGeneral);
			cellStyleTitle.setFont(createFontArial(workbook, 10));
			cellStyleTitle.setAlignment(HorizontalAlignment.CENTER);
			styleMap.put(CellStyle.CELL_STYLE_TYPE_TITLE.getCode(), cellStyleTitle);

			org.apache.poi.ss.usermodel.CellStyle cellStyleColumnHeader = workbook.createCellStyle();
			cellStyleColumnHeader.cloneStyleFrom(cellStyleGeneral);
			cellStyleColumnHeader.setFont(createFontArial(workbook, 10));
			cellStyleColumnHeader.setAlignment(HorizontalAlignment.CENTER);
			styleMap.put(CellStyle.CELL_STYLE_TYPE_COLUMN_HEADER.getCode(), cellStyleColumnHeader);

			org.apache.poi.ss.usermodel.CellStyle cellStyleRowHeader = workbook.createCellStyle();
			cellStyleRowHeader.cloneStyleFrom(cellStyleGeneral);
			cellStyleRowHeader.setFont(createFontArial(workbook, 10));
			cellStyleRowHeader.setAlignment(HorizontalAlignment.CENTER);
			cellStyleRowHeader.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
			cellStyleRowHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			styleMap.put(CellStyle.CELL_STYLE_TYPE_ROW_HEADER.getCode(), cellStyleRowHeader);

			org.apache.poi.ss.usermodel.CellStyle cellStyleGrayRowHeader = workbook.createCellStyle();
			cellStyleGrayRowHeader.cloneStyleFrom(cellStyleRowHeader);
			cellStyleGrayRowHeader.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			styleMap.put(CellStyle.CELL_STYLE_TYPE_ROW_HEADER_GREY.getCode(), cellStyleGrayRowHeader);

			org.apache.poi.ss.usermodel.CellStyle cellStyleRowHeaderTopAlign = workbook.createCellStyle();
			cellStyleRowHeaderTopAlign.cloneStyleFrom(cellStyleRowHeader);
			cellStyleRowHeaderTopAlign.setVerticalAlignment(VerticalAlignment.TOP);
			styleMap.put(CellStyle.CELL_STYLE_TYPE_ROW_HEADER_TOP_ALIGN.getCode(), cellStyleRowHeaderTopAlign);

			org.apache.poi.ss.usermodel.CellStyle cellStyleHideHeader = workbook.createCellStyle();
			cellStyleHideHeader.setFont(createFontWhiteArial(workbook, 10));
			cellStyleHideHeader.setWrapText(true);
			cellStyleHideHeader.setAlignment(HorizontalAlignment.LEFT);
			cellStyleHideHeader.setVerticalAlignment(VerticalAlignment.CENTER);
			cellStyleHideHeader.setLocked(true);
			cellStyleHideHeader.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			styleMap.put(CellStyle.CELL_STYLE_TYPE_HEADER_HIDE.getCode(), cellStyleHideHeader);

			org.apache.poi.ss.usermodel.CellStyle cellStyleWhiteHideHeader = workbook.createCellStyle();
			cellStyleWhiteHideHeader.cloneStyleFrom(cellStyleHideHeader);
			cellStyleWhiteHideHeader.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			styleMap.put(CellStyle.CELL_STYLE_TYPE_HEADER_HIDE_WHITE.getCode(), cellStyleWhiteHideHeader);

			org.apache.poi.ss.usermodel.CellStyle cellStyleValue = workbook.createCellStyle();
			cellStyleValue.cloneStyleFrom(cellStyleGeneral);
			cellStyleValue.setFont(createFontArial(workbook, 10));
			cellStyleValue.setAlignment(HorizontalAlignment.CENTER);
			cellStyleValue.setWrapText(true);
			cellStyleValue.setLocked(false);
			styleMap.put(CellStyle.CELL_STYLE_TYPE_VALUE.getCode(), cellStyleValue);

			cellStyleCache.put(workbook, styleMap);
		}
		return cellStyleCache.get(workbook).get(cellStyleType);
	}

	/**
	 * Merge cells begin from {firstRow}, {lastRow} to {firstColumn}, {lastColumn}.
	 *
	 * @param sheet       given sheet
	 * @param firstRow    first row index
	 * @param lastRow     last row index
	 * @param firstColumn first column index
	 * @param lastColumn  last column index
	 */
	public static boolean mergeCells(Sheet sheet, int firstRow, int lastRow, int firstColumn, int lastColumn) {
		if (sheet == null) {
			logger.warn("Sheet is null!");
			return false;
		}
		CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
		sheet.addMergedRegion(cellRangeAddress);
		return true;
	}

	/**
	 * Set default row height of the sheet.
	 *
	 * @param sheet  given sheet
	 * @param height height
	 * @return true/false
	 */
	public static boolean setDefaultRowHeight(Sheet sheet, int height) {
		if (sheet == null) {
			logger.warn("Sheet is null!");
			return false;
		}
		try {
			sheet.setDefaultRowHeight(Short.parseShort(String.valueOf(height)));
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
		return true;
	}

	/**
	 * Set height of row. Row should be exists.
	 * <pre>
	 *     NOTE:invoke after row is created.
	 * </pre>
	 *
	 * @param sheet  given sheet
	 * @param row    row number, begin from 0.
	 * @param height height
	 * @return true/false
	 */
	public static boolean setRowHeight(Sheet sheet, int row, int height) {
		if (sheet == null) {
			logger.warn("Sheet is null!");
			return false;
		}
		try {
			Row r = sheet.getRow(row);
			if (r == null) {
				logger.warn(String.format("Row for number:[%s] is null!", row));
				return false;
			}
			r.setHeight(Short.parseShort(String.valueOf(height)));
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
		return true;
	}

	/**
	 * Hide row, Set height 0 of row. Row should be exists.
	 * <pre>
	 *     NOTE:invoke after row is created.
	 * </pre>
	 *
	 * @param sheet given sheet
	 * @param row   row number, begin from 0.
	 * @return true/false
	 */
	public static boolean hideRow(Sheet sheet, int row) {
		if (sheet == null) {
			logger.warn("Sheet is null!");
			return false;
		}
		try {
			Row r = sheet.getRow(row);
			if (r == null) {
				logger.warn(String.format("Row for number:[%s] is null!", row));
				return false;
			}
			r.setZeroHeight(true);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
		return true;
	}

	/**
	 * Set width of column.
	 *
	 * @param sheet  given sheet
	 * @param column column number, begin from 0.
	 * @param width  width
	 * @return true/false
	 */
	public static boolean setColumnWidth(Sheet sheet, int column, int width) {
		if (sheet == null) {
			logger.warn("Sheet is null!");
			return false;
		}
		try {
			sheet.setColumnWidth(column, width);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
		return true;
	}

	/**
	 * Hide column, Set width 0 of column.
	 *
	 * @param sheet  given sheet
	 * @param column column number, begin from 0.
	 * @return true/false
	 */
	public static boolean hideColumn(Sheet sheet, int column) {
		if (sheet == null) {
			logger.warn("Sheet is null!");
			return false;
		}
		try {
			sheet.setColumnHidden(column, true);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
		return true;
	}

	/**
	 * insert value into cell
	 *
	 * @param sheet     sheet
	 * @param column    column
	 * @param row       row
	 * @param value     value
	 * @param cellStyle cell style
	 */
	public static void insertCellValue(Sheet sheet, int column, int row, String value,
			org.apache.poi.ss.usermodel.CellStyle cellStyle) {
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
	 * get value of cell
	 *
	 * @param sheet  sheet
	 * @param column column
	 * @param row    row
	 * @return cell value
	 */
	public static String getCellValue(Sheet sheet, int column, int row) {
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
				switch (cell.getCellTypeEnum()) {
				case STRING:
					cellValue = cell.getRichStringCellValue().getString().trim();
					break;
				case NUMERIC:
					cellValue = decimalFormat.format(cell.getNumericCellValue()).trim();
					break;
				case BOOLEAN:
					cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
					break;
				case FORMULA:
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
	 * Get sheet at sheetNumber from a workbook, create with sheetName if not exists.
	 *
	 * @param workbook    given workbook
	 * @param sheetNumber given sheet number
	 * @param sheetName   given sheet name (can be null)
	 * @return sheet
	 */
	public static Sheet getSheetCreateIfNotExist(Workbook workbook, int sheetNumber, String sheetName) {
		if (workbook == null)
			return null;
		if (sheetNumber < 0)
			return null;
		Sheet sheet;
		try {
			sheet = workbook.getSheetAt(sheetNumber);
		} catch (Exception e) {
			String name = sheetName;
			if (name == null || name.isEmpty()) {
				name = String.valueOf(sheetNumber);
			}
			sheet = workbook.createSheet(name);
			workbook.setSheetOrder(name, sheetNumber);
		}
		return sheet;
	}

	/**
	 * write workbook into disk file
	 * <p>
	 * <pre>
	 *     NOTE : Existing file can not be overwrite.
	 *     If the workbook is opened from the file which is exist,
	 *     you should create a new file to write in, then close the
	 *     existing file which the workbook opened from and delete
	 *     it, finally rename the new file with the same name.
	 *
	 *     {@link org.apache.poi.POIXMLDocument#write(OutputStream)}
	 * </pre>
	 *
	 * @param workbook excel workbook
	 * @param file     disk file
	 * @return true if succeeded, false if failed
	 */
	public static boolean writeFileToDisk(Workbook workbook, File file) {
		if (file == null || !file.exists())
			return false;
		if (workbook == null)
			return false;
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream(file);
			workbook.write(fileOutputStream);
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		} finally {
			if (fileOutputStream != null) {
				try {
					fileOutputStream.close();
				} catch (Exception e) {
					logger.warn(e.getMessage(), e);
				}
			}
		}
		return true;
	}

	/**
	 * Close the workbook
	 *
	 * @param workbook excel workbook
	 * @return true/false
	 */
	public static boolean closeWorkbook(Workbook workbook) {
		if (workbook == null)
			return false;
		try {
			workbook.close();
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			return false;
		}
		cellStyleCache.remove(workbook);
		return true;
	}
}
