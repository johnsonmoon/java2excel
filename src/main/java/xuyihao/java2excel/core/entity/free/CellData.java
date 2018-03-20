package xuyihao.java2excel.core.entity.free;

import xuyihao.java2excel.core.entity.Common.CellStyle;

/**
 * Cell data entity.
 * <p>
 * Created by xuyh at 2018/3/20 15:05.
 */
public class CellData {
	/**
	 * row location start from 0
	 */
	private int row;
	/**
	 * column location start from 0
	 */
	private int column;
	/**
	 * data
	 */
	private Object data;
	/**
	 * cell style format {@link CellStyle}
	 */
	private CellStyle cellStyle;

	public CellData() {
	}

	/**
	 * Create a cellData instance.
	 *
	 * @param row       row location start from 0
	 * @param column    column location start from 0
	 * @param data      data
	 * @param cellStyle cell style format {@link CellStyle}
	 */
	public CellData(int row, int column, Object data, CellStyle cellStyle) {
		this.row = row;
		this.column = column;
		this.data = data;
		this.cellStyle = cellStyle;
	}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getColumn() {
		return column;
	}

	public void setColumn(int column) {
		this.column = column;
	}

	public Object getData() {
		return data;
	}

	public void setData(Object data) {
		this.data = data;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	@Override
	public String toString() {
		return "CellData{" +
				"row=" + row +
				", column=" + column +
				", data=" + data +
				", cellStyle=" + cellStyle +
				'}';
	}
}
