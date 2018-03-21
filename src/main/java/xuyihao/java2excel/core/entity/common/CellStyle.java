package xuyihao.java2excel.core.entity.common;

/**
 * Created by xuyh at 2018/3/20 15:22.
 */
public enum CellStyle {
	CELL_STYLE_TYPE_GENERAL(0), 
	CELL_STYLE_TYPE_TITLE(1), 
	CELL_STYLE_TYPE_COLUMN_HEADER(2), 
	CELL_STYLE_TYPE_ROW_HEADER(3), 
	CELL_STYLE_TYPE_ROW_HEADER_GREY(4),
	CELL_STYLE_TYPE_ROW_HEADER_TOP_ALIGN(5), 
	CELL_STYLE_TYPE_HEADER_HIDE(6), 
	CELL_STYLE_TYPE_HEADER_HIDE_WHITE(7), 
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
