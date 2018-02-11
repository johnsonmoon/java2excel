package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.entity.custom.map.ColumnMapper;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.core.operation.custom.CustomExporter;
import xuyihao.java2excel.util.AnnotationUtils;
import xuyihao.java2excel.util.FileUtils;
import xuyihao.java2excel.util.ReflectionUtils;
import xuyihao.java2excel.util.ValueUtils;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by xuyh at 2018/2/11 17:39.
 */
public abstract class CustomAbstractEdior {
	/**
	 * Generate excel column mapping instance from given class.
	 * <p>
	 * <pre>
	 *     If given class fields has not declared @Column annotation,
	 *     generate from reflection only. Otherwise generate by annotation declaration.
	 * </pre>
	 *
	 * @param tClass given class
	 * @see xuyihao.java2excel.core.entity.custom.map.annotation.Column
	 */
	public ColumnMapper generateColumnMapper(Class<?> tClass) {
		if (!AnnotationUtils.hasAnnotationColumn(tClass.getDeclaredFields())) {
			return generateColumnMapperWithReflection(tClass);
		} else {
			return generateColumnMapperWithAnnotation(tClass);
		}
	}

	private ColumnMapper generateColumnMapperWithReflection(Class<?> clazz) {
		ColumnMapper columnMapper = new ColumnMapper();
		int columnNumber = 0;
		for (Field field : ReflectionUtils.getFieldsUnStaticUnFinal(clazz)) {
			String fieldName = field.getName();
			columnMapper.add(columnNumber, fieldName);
			columnNumber++;
		}
		return columnMapper;
	}

	private ColumnMapper generateColumnMapperWithAnnotation(Class<?> clazz) {
		Field[] fields = clazz.getDeclaredFields();
		if (!AnnotationUtils.hasAnnotationColumn(fields))
			return null;
		ColumnMapper columnMapper = new ColumnMapper();
		for (Field field : fields) {
			if (AnnotationUtils.hasAnnotationColumn(field)) {
				int columnNumber = AnnotationUtils.getAnnotationColumn(field).column();
				String fieldName = field.getName();
				columnMapper.add(columnNumber, fieldName);
			}
		}
		return columnMapper;
	}

	/**
	 * Generate data map list from given type instance list.
	 * <p>
	 * <pre>
	 * 	Map list value. Complex data(map, list .etc) with json format.(For read data from excel into objects).
	 * </pre>
	 *
	 * @param tList        given type instance list.
	 * @param columnMapper column mapping information
	 * @return model data list
	 */
	private List<Map<String, Object>> generateData(List<?> tList, ColumnMapper columnMapper) {
		List<Map<String, Object>> mapList = new ArrayList<>();
		if (tList == null || tList.isEmpty())
			return mapList;
		for (Object t : tList) {
			Map<String, Object> map = new HashMap<>();
			for (ColumnMapper.Map columnMap : columnMapper.getOrderedMaps(ColumnMapper.ORDER_TYPE_ASC)) {
				String fieldName = columnMap.fieldName;
				Object value = ReflectionUtils.getFieldValue(t, fieldName);
				map.put(fieldName, ValueUtils.formatValue(value));
			}
			mapList.add(map);
		}
		return mapList;
	}

	/**
	 * @param columnMapper given column mapping information
	 * @param workbook     given workbook
	 * @param sheetNumber  given sheet number
	 * @param sheetName    given start row number to write with
	 * @return true/false
	 */
	protected boolean writeColumnMapperHeader(ColumnMapper columnMapper, Workbook workbook, int sheetNumber,
			String sheetName) {
		return CustomExporter.createExcel(workbook, sheetNumber, sheetName, columnMapper);
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param dataMapList    given data map list
	 * @param workbook       given workbook
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	protected boolean writeDataDirectly(List<Map<Integer, Object>> dataMapList, Workbook workbook, int sheetNumber,
			int startRowNumber) {
		if (dataMapList == null || dataMapList.isEmpty())
			return false;
		return CustomExporter.insertExcelData(workbook, sheetNumber, startRowNumber, dataMapList);
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param dataMapList    given data map list
	 * @param columnMapper   given column mapping information
	 * @param workbook       given workbook
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	protected boolean writeDataDirectly(List<Map<String, Object>> dataMapList, ColumnMapper columnMapper,
			Workbook workbook, int sheetNumber, int startRowNumber) {
		if (dataMapList == null || dataMapList.isEmpty())
			return false;
		return CustomExporter.insertExcelData(workbook, sheetNumber, startRowNumber, columnMapper, dataMapList);
	}

	/**
	 * Write data into a workbook's given sheet starting from given start row number.
	 *
	 * @param tList          given clazz data list
	 * @param workbook       given workbook
	 * @param sheetNumber    given sheet number
	 * @param startRowNumber given start row number to write with
	 * @return true/false
	 */
	protected boolean writeData(List<?> tList, Workbook workbook, int sheetNumber, int startRowNumber) {
		if (tList == null || tList.isEmpty())
			return false;
		ColumnMapper columnMapper = generateColumnMapper(tList.get(0).getClass());
		List<Map<String, Object>> dataMapList = generateData(tList, columnMapper);
		return CustomExporter.insertExcelData(workbook, sheetNumber, startRowNumber, columnMapper, dataMapList);
	}

	/**
	 * Write existing workbook into a given file. Create a new file if the file isn't exist.
	 *
	 * @param workbook     given workbook
	 * @param filePathName given file path name
	 * @return true/false
	 * @throws Exception exceptions
	 */
	protected boolean flush(Workbook workbook, String filePathName) throws Exception {
		if (workbook == null)
			throw new NullPointerException("workbook is null!");
		File file = FileUtils.createFile(filePathName);
		return Common.writeFileToDisk(workbook, file);
	}
}
