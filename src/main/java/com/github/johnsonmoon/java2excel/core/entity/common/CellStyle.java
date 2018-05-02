package com.github.johnsonmoon.java2excel.core.entity.common;

/**
 * Created by xuyh at 2018/3/20 15:22.
 */
public enum CellStyle {
	/**
	 * General style.
	 */
	CELL_STYLE_TYPE_GENERAL(0),
	/**
	 * Title style.
	 */
	CELL_STYLE_TYPE_TITLE(1),
	/**
	 * Column header style.
	 */
	CELL_STYLE_TYPE_COLUMN_HEADER(2),
	/**
	 * Header row style.
	 */
	CELL_STYLE_TYPE_ROW_HEADER(3),
	/**
	 * Grey header row style.
	 */
	CELL_STYLE_TYPE_ROW_HEADER_GREY(4),
	/**
	 * Top row header align style.
	 */
	CELL_STYLE_TYPE_ROW_HEADER_TOP_ALIGN(5),
	/**
	 * Header hide style.
	 */
	CELL_STYLE_TYPE_HEADER_HIDE(6),
	/**
	 * White header hide style.
	 */
	CELL_STYLE_TYPE_HEADER_HIDE_WHITE(7),
	/**
	 * Value style.
	 */
	CELL_STYLE_TYPE_VALUE(8);

	private int code;

	CellStyle(int code) {
		this.code = code;
	}

	public int getCode() {
		return code;
	}

	@Override
	public String toString() {
		return "CellStyle{" +
				"code=" + code +
				'}';
	}
}
