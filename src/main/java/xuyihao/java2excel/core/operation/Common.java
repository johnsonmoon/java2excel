package xuyihao.java2excel.core.operation;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
	 * cell style cache for caching cell styles
	 * <pre>
	 *     We can define only up to 64000 style in a .xlsx Workbook.
	 *     So cell styles for workbook should be cached.
	 *     After workbook has been closed, cache needed to be cleaned.
	 * </pre>
	 */
	private static Map<Workbook, Map<Integer, CellStyle>> cellStyleCache = new HashMap<>();

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
	public static CellStyle createCellStyle(Workbook workbook, int cellStyleType) {
		if (!cellStyleCache.containsKey(workbook)) {
			Map<Integer, CellStyle> styleMap = new HashMap<>();

			CellStyle cellStyleGeneral = workbook.createCellStyle();
			cellStyleGeneral.setWrapText(false);
			cellStyleGeneral.setVerticalAlignment(VerticalAlignment.CENTER);
			styleMap.put(CELL_STYLE_TYPE_GENERAL, cellStyleGeneral);

			CellStyle cellStyleTitle = workbook.createCellStyle();
			cellStyleTitle.cloneStyleFrom(cellStyleGeneral);
			cellStyleTitle.setFont(createFontArial(workbook, 10));
			cellStyleTitle.setAlignment(HorizontalAlignment.CENTER);
			styleMap.put(CELL_STYLE_TYPE_TITLE, cellStyleTitle);

			CellStyle cellStyleColumnHeader = workbook.createCellStyle();
			cellStyleColumnHeader.cloneStyleFrom(cellStyleGeneral);
			cellStyleColumnHeader.setFont(createFontArial(workbook, 10));
			cellStyleColumnHeader.setAlignment(HorizontalAlignment.CENTER);
			styleMap.put(CELL_STYLE_TYPE_COLUMN_HEADER, cellStyleColumnHeader);

			CellStyle cellStyleRowHeader = workbook.createCellStyle();
			cellStyleRowHeader.cloneStyleFrom(cellStyleGeneral);
			cellStyleRowHeader.setFont(createFontArial(workbook, 10));
			cellStyleRowHeader.setAlignment(HorizontalAlignment.CENTER);
			cellStyleRowHeader.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
			cellStyleRowHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			styleMap.put(CELL_STYLE_TYPE_ROW_HEADER, cellStyleRowHeader);

			CellStyle cellStyleGrayRowHeader = workbook.createCellStyle();
			cellStyleGrayRowHeader.cloneStyleFrom(cellStyleRowHeader);
			styleMap.put(CELL_STYLE_TYPE_ROW_HEADER_GRAY, cellStyleGrayRowHeader);

			CellStyle cellStyleRowHeaderTopAlign = workbook.createCellStyle();
			cellStyleRowHeaderTopAlign.cloneStyleFrom(cellStyleRowHeader);
			cellStyleRowHeaderTopAlign.setVerticalAlignment(VerticalAlignment.TOP);
			styleMap.put(CELL_STYLE_TYPE_ROW_HEADER_TOP_ALIGN, cellStyleRowHeaderTopAlign);

			CellStyle cellStyleHideHeader = workbook.createCellStyle();
			cellStyleHideHeader.setFont(createFontWhiteArial(workbook, 10));
			cellStyleHideHeader.setWrapText(true);
			cellStyleHideHeader.setAlignment(HorizontalAlignment.LEFT);
			cellStyleHideHeader.setVerticalAlignment(VerticalAlignment.CENTER);
			cellStyleHideHeader.setLocked(true);
			cellStyleHideHeader.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			styleMap.put(CELL_STYLE_TYPE_HEADER_HIDE, cellStyleHideHeader);

			CellStyle cellStyleWhiteHideHeader = workbook.createCellStyle();
			cellStyleWhiteHideHeader.cloneStyleFrom(cellStyleHideHeader);
			cellStyleWhiteHideHeader.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			styleMap.put(CELL_STYLE_TYPE_HEADER_HIDE_WHITE, cellStyleWhiteHideHeader);

			CellStyle cellStyleValue = workbook.createCellStyle();
			cellStyleValue.cloneStyleFrom(cellStyleGeneral);
			cellStyleValue.setFont(createFontArial(workbook, 10));
			cellStyleValue.setAlignment(HorizontalAlignment.CENTER);
			cellStyleValue.setWrapText(true);
			cellStyleValue.setLocked(false);
			styleMap.put(CELL_STYLE_TYPE_VALUE, cellStyleValue);

			cellStyleCache.put(workbook, styleMap);
		}
		return cellStyleCache.get(workbook).get(cellStyleType);
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
	 * @throws Exception exceptions
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
