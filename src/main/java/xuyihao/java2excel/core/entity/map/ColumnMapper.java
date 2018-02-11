package xuyihao.java2excel.core.entity.map;

import java.util.ArrayList;
import java.util.List;

/**
 * Mapper for [fieldName:column] mapping.
 * <p>
 * Created by xuyh at 2018/2/10 10:41.
 */
public class ColumnMapper {
	/**
	 * order type desc
	 */
	public static final String ORDER_TYPE_DESC = "desc";
	/**
	 * order type asc
	 */
	public static final String ORDER_TYPE_ASC = "asc";

	private List<Map> maps = new ArrayList<>();

	/**
	 * Add mapping.
	 *
	 * @param map {@link ColumnMapper.Map}
	 * @return true/false
	 */
	public boolean add(Map map) {
		if (map == null)
			return false;
		if (map.fieldName == null || map.columnNumber == null)
			return false;
		if (containsColumnNumber(map.columnNumber))
			return false;
		if (containsFieldName(map.fieldName))
			return false;
		return this.maps.add(map);
	}

	/**
	 * Add mapping.
	 *
	 * @param columnNumber column number (start from 0.)
	 * @param fieldName    field name
	 * @return true/false
	 */
	public boolean add(Integer columnNumber, String fieldName) {
		if (columnNumber == null || fieldName == null)
			return false;
		Map map = new Map();
		map.columnNumber = columnNumber;
		map.fieldName = fieldName;
		return this.maps.add(map);
	}

	/**
	 * Whether contains column number.
	 *
	 * @param columnNumber column number
	 * @return true/false
	 */
	public boolean containsColumnNumber(Integer columnNumber) {
		for (Map map : this.maps) {
			if (map.columnNumber.equals(columnNumber))
				return true;
		}
		return false;
	}

	/**
	 * Whether contains field name.
	 *
	 * @param fieldName field name
	 * @return true/false
	 */
	public boolean containsFieldName(String fieldName) {
		for (Map map : this.maps) {
			if (map.fieldName.equals(fieldName))
				return true;
		}
		return false;
	}

	/**
	 * Get column number list in mapping.
	 *
	 * @return column number list
	 */
	public List<Integer> getColumnNumbers() {
		List<Integer> list = new ArrayList<>();
		for (Map map : this.maps) {
			list.add(map.columnNumber);
		}
		return list;
	}

	/**
	 * Get field name list in mapping.
	 *
	 * @return field name list
	 */
	public List<String> getFieldNames() {
		List<String> list = new ArrayList<>();
		for (Map map : this.maps) {
			list.add(map.fieldName);
		}
		return list;
	}

	/**
	 * Get mapping ordered by column number;
	 *
	 * @param orderType ordering type
	 *                  {@link ColumnMapper#ORDER_TYPE_ASC}
	 *                  {@link ColumnMapper#ORDER_TYPE_DESC}
	 * @return {@link ColumnMapper.Map}
	 */
	public List<Map> getOrderedMaps(String orderType) {
		List<Map> list = new ArrayList<>();
		for (Map m : this.maps) {
			Map map = new Map();
			map.columnNumber = m.columnNumber;
			map.fieldName = m.fieldName;
			list.add(map);
		}
		list.sort((map1, map2) -> {
			if (map1.columnNumber > map2.columnNumber) {
				if (orderType.equals(ORDER_TYPE_DESC))
					return -1;
				if (orderType.equals(ORDER_TYPE_ASC))
					return 1;
				return 1;
			} else if (map1.columnNumber < map2.columnNumber) {
				if (orderType.equals(ORDER_TYPE_DESC))
					return 1;
				if (orderType.equals(ORDER_TYPE_ASC))
					return -1;
				return -1;
			} else {
				return 0;
			}
		});
		return list;
	}

	/**
	 * [columnNumber:fieldName] mapping
	 */
	public static class Map {
		/**
		 * column number, begin from 0.
		 */
		public Integer columnNumber;
		/**
		 * field name
		 */
		public String fieldName;

		@Override
		public String toString() {
			return "Map{" +
					"columnNumber=" + columnNumber +
					", fieldName='" + fieldName + '\'' +
					'}';
		}
	}

	@Override
	public String toString() {
		return "ColumnMapper{" +
				"maps=" + maps +
				'}';
	}
}
