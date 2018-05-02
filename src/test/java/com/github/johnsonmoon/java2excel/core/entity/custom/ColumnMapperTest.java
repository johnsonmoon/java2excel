package com.github.johnsonmoon.java2excel.core.entity.custom;

import org.junit.Assert;
import org.junit.Test;

import java.util.List;

/**
 * Created by xuyh at 2018/2/10 11:39.
 */
public class ColumnMapperTest {
	@Test
	public void test() {
		ColumnMapper columnMapper = new ColumnMapper();
		ColumnMapper.Map map = new ColumnMapper.Map();
		map.fieldName = "number";
		map.columnNumber = 0;
		Assert.assertTrue(columnMapper.add(map));
		Assert.assertTrue(columnMapper.add(1, "name"));
		Assert.assertTrue(columnMapper.add(2, "phone"));
		Assert.assertTrue(columnMapper.add(3, "email"));
		Assert.assertTrue(columnMapper.containsColumnNumber(1));
		Assert.assertTrue(columnMapper.containsFieldName("name"));
		List<String> fields = columnMapper.getFieldNames();
		Assert.assertNotNull(fields);
		fields.forEach(System.out::println);
		List<Integer> columns = columnMapper.getColumnNumbers();
		Assert.assertNotNull(columns);
		columns.forEach(System.out::println);
		List<ColumnMapper.Map> orderedMapping = columnMapper.getOrderedMaps(ColumnMapper.ORDER_TYPE_DESC);
		Assert.assertNotNull(orderedMapping);
		orderedMapping.forEach(mapQ -> {
			System.out.println(mapQ.columnNumber + "  " + mapQ.fieldName);
		});
	}
}
