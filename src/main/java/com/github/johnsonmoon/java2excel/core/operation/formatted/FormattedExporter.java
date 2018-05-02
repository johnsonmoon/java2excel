package com.github.johnsonmoon.java2excel.core.operation.formatted;

import com.github.johnsonmoon.java2excel.core.entity.common.CellStyle;
import com.github.johnsonmoon.java2excel.core.entity.formatted.dict.Meta;
import com.github.johnsonmoon.java2excel.core.entity.formatted.model.Model;
import com.github.johnsonmoon.java2excel.core.operation.Common;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import com.github.johnsonmoon.java2excel.util.StringUtils;

import java.util.*;

/**
 * Created by Xuyh at 2016/07/22 上午 11:36.
 */
public class FormattedExporter {
	private static Logger logger = LoggerFactory.getLogger(FormattedExporter.class);

	/**
	 * create excel sheet
	 *
	 * @param workbook excel workbook
	 * @param sheetNum sheet number
	 * @param model    model
	 * @param language meta info language(en_US, zh_CN, etc.)
	 *                 {@link Meta}
	 * @return true/false
	 */
	public static boolean createExcel(final Workbook workbook, int sheetNum, Model model, String language) {
		boolean flag;
		try {
			if (workbook == null)
				return false;
			if (model == null)
				return false;
			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			// column summary
			int columnSize;
			int attrValueSize = model.getAttributes().size();
			columnSize = attrValueSize + 1;
			// column width
			for (int i = 0; i < columnSize; i++) {
				sheet.setColumnWidth(i + 1, 4800);
			}
			// hide column
			sheet.setColumnHidden(1, true);
			// hide row
			sheet.createRow(3).setZeroHeight(true);
			// merge cells
			try {
				CellRangeAddress cellRangeAddress1 = new CellRangeAddress(0, 1, 0, 0);
				CellRangeAddress cellRangeAddress3 = new CellRangeAddress(0, 1, 1, columnSize);
				sheet.addMergedRegion(cellRangeAddress1);
				sheet.addMergedRegion(cellRangeAddress3);
			} catch (Exception e) {
				logger.warn(e.getMessage(), e);
			}
			// insert meta info
			Common.insertCellValue(sheet, 0, 0, model.getJavaClassName(),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_HEADER_HIDE.getCode()));
			Common.insertCellValue(sheet, 1, 0, model.getName(),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_TITLE.getCode()));
			Common.insertCellValue(sheet, 0, 2, Meta.FILEDS.getMeta(language),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER.getCode()));
			Common.insertCellValue(sheet, 0, 3, Meta.DONT_MODIFY.getMeta(language),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER_GREY
									.getCode()));
			Common.insertCellValue(sheet, 0, 4, Meta.DATA_FORMAT.getMeta(language),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER.getCode()));
			Common.insertCellValue(sheet, 0, 5, Meta.DEFAULT_VALUE.getMeta(language),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER.getCode()));
			Common.insertCellValue(sheet, 0, 6, Meta.DATA.getMeta(language),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER_TOP_ALIGN
									.getCode()));
			Common.insertCellValue(sheet, 1, 2, Meta.DATA_FLAG.getMeta(language),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_COLUMN_HEADER.getCode()));
			Common.insertCellValue(sheet, 1, 3, Meta.DATA_FLAG.getMeta(language),
					Common.createCellStyle(workbook,
							com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_HEADER_HIDE_WHITE
									.getCode()));
			for (int i = 0; i < model.getAttributes().size(); i++) {
				String label;
				if (model.getAttributes().get(i).getUnit() == null
						|| model.getAttributes().get(i).getUnit().equals("")) {
					label = model.getAttributes().get(i).getAttrName();
				} else {
					label = model.getAttributes().get(i).getAttrName() + "("
							+ model.getAttributes().get(i).getUnit() + ")";
				}
				String attrInfo = model.getAttributes().get(i).getAttrCode()
						+ "&&" + model.getAttributes().get(i).getAttrName()
						+ "&&" + model.getAttributes().get(i).getAttrType()
						+ "&&" + model.getAttributes().get(i).getFormatInfo()
						+ "&&" + model.getAttributes().get(i).getDefaultValue()
						+ "&&" + model.getAttributes().get(i).getUnit()
						+ "&&" + model.getAttributes().get(i).getJavaClassName();
				String formatInfo = model.getAttributes().get(i).getFormatInfo();
				String defaultValue = model.getAttributes().get(i).getDefaultValue();
				Common.insertCellValue(sheet, i + 2, 2, label,
						Common.createCellStyle(workbook,
								com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER.getCode()));
				Common.insertCellValue(sheet, i + 2, 3, attrInfo,
						Common.createCellStyle(workbook,
								com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER_GREY
										.getCode()));
				Common.insertCellValue(sheet, i + 2, 4, formatInfo,
						Common.createCellStyle(workbook,
								com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER.getCode()));
				Common.insertCellValue(sheet, i + 2, 5, defaultValue,
						Common.createCellStyle(workbook,
								com.github.johnsonmoon.java2excel.core.entity.common.CellStyle.CELL_STYLE_TYPE_ROW_HEADER.getCode()));
			}
			flag = true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}

	/**
	 * insert data
	 *
	 * @param workbook    excel workbook
	 * @param sheetNum    sheet number
	 * @param startRowNum row number to start writing (begin from 0)
	 * @param modelList   model data list to write with
	 * @return true/false
	 */
	public static boolean insertExcelData(final Workbook workbook, int sheetNum, int startRowNum,
			List<Model> modelList) {
		boolean flag;
		try {
			if (modelList == null || modelList.isEmpty())
				return false;

			if (workbook == null) {
				return false;
			}

			Sheet sheet = Common.getSheetCreateIfNotExist(workbook, sheetNum, null);
			if (sheet == null) {
				return false;
			}

			String identifyString = Common.getCellValue(sheet, 0, 0);
			if (!identifyString.equals(StringUtils.replaceNullToEmpty(modelList.get(0).getJavaClassName()))) {
				return false;
			}

			if (startRowNum < 0) {
				logger.warn("Data insert must above 0!");
			}
			for (int j = 0; j < modelList.size(); j++) {
				for (int i = 0; i < modelList.get(j).getAttrValues().size(); i++) {
					Common.insertCellValue(sheet, i + 2, j + startRowNum + 6, modelList.get(j).getAttrValues().get(i),
							Common.createCellStyle(workbook, CellStyle.CELL_STYLE_TYPE_VALUE.getCode()));
				}
			}
			flag = true;
		} catch (Exception e) {
			logger.warn(e.getMessage(), e);
			flag = false;
		}
		return flag;
	}
}
