package xuyihao.java2excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import xuyihao.java2excel.core.entity.custom.map.ColumnMapper;
import xuyihao.java2excel.core.operation.Common;
import xuyihao.java2excel.core.operation.custom.CustomExporter;
import xuyihao.java2excel.core.operation.custom.CustomImporter;
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
@SuppressWarnings("all")
public abstract class CustomAbstractEditor {
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
    public List<Map<String, Object>> generateData(List<?> tList, ColumnMapper columnMapper) {
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
     * @param claszz      type
     * @param workbook    given workbook
     * @param sheetNumber given sheet number
     * @param sheetName   given start row number to write with
     * @return true/false
     */
    protected boolean writeColumnMapperHeader(Class<?> claszz, Workbook workbook, int sheetNumber,
                                              String sheetName) {
        ColumnMapper columnMapper = generateColumnMapper(claszz);
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
                                        int startRowNumber) throws Exception {
        if (dataMapList == null || dataMapList.isEmpty())
            throw new RuntimeException("dataMapList is empty or null.");
        if (workbook == null)
            throw new NullPointerException("workbook is null.");
        if (sheetNumber < 0 || startRowNumber < 0)
            throw new RuntimeException("shettNumber or startRowNumber can not be less than 0.");
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
                                        Workbook workbook, int sheetNumber, int startRowNumber) throws Exception {
        if (dataMapList == null || dataMapList.isEmpty()) {
            throw new RuntimeException("dataMapList is empty or null.");
        }
        if (columnMapper == null) {
            throw new NullPointerException("columnMapper is null.");
        }
        if (workbook == null) {
            throw new NullPointerException("workbook is null.");
        }
        if (sheetNumber < 0 || startRowNumber < 0) {
            throw new RuntimeException("shettNumber or startRowNumber can not be less than 0.");
        }
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
        if (tList == null || tList.isEmpty()) {
            throw new RuntimeException("tList is empty or null.");
        }
        if (workbook == null) {
            throw new NullPointerException("workbook is null.");
        }
        if (sheetNumber < 0 || startRowNumber < 0) {
            throw new RuntimeException("shettNumber or startRowNumber can not be less than 0.");
        }
        ColumnMapper columnMapper = generateColumnMapper(tList.get(0).getClass());
        List<Map<String, Object>> dataMapList = generateData(tList, columnMapper);
        return CustomExporter.insertExcelData(workbook, sheetNumber, startRowNumber, columnMapper, dataMapList);
    }

    /**
     * Write data into a workbook's given sheet starting from given start row number.
     *
     * @param tList          given clazz data list
     * @param columnMapper   column mapping information
     * @param workbook       given workbook
     * @param sheetNumber    given sheet number
     * @param startRowNumber given start row number to write with
     * @return true/false
     */
    protected boolean writeData(List<?> tList, ColumnMapper columnMapper, Workbook workbook, int sheetNumber, int startRowNumber) {
        if (tList == null || tList.isEmpty()) {
            throw new RuntimeException("tList is empty or null.");
        }
        if (columnMapper == null)
            throw new NullPointerException("columnMapper is null.");
        if (workbook == null) {
            throw new NullPointerException("workbook is null.");
        }
        if (sheetNumber < 0 || startRowNumber < 0) {
            throw new RuntimeException("shettNumber or startRowNumber can not be less than 0.");
        }
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

    /**
     * Read map data from a given workbook in sheet from a sheetNumber.
     *
     * @param workbook       given workbook
     * @param sheetNumber    given sheet number
     * @param beginRowNumber row number to begin reading
     * @param readSize       row count to read
     * @return map list
     * @throws Exception exceptions
     */
    protected List<Map<Integer, Object>> readDataDirectly(Workbook workbook, int sheetNumber, int beginRowNumber,
                                                          int readSize)
            throws Exception {
        if (workbook == null)
            throw new NullPointerException("workbook can not be null.");
        if (sheetNumber < 0)
            throw new RuntimeException("sheet number can not below 0.");
        if (beginRowNumber < 0)
            throw new RuntimeException("beginRowNumber can not below 0.");
        if (readSize < 0)
            throw new RuntimeException("readSize can not below 0.");
        List<Map<Integer, Object>> mapList = new ArrayList<>();
        mapList.addAll(CustomImporter.getDataFromExcel(workbook, sheetNumber, beginRowNumber, readSize));
        return mapList;
    }

    /**
     * Read map data from a given workbook in sheet from a sheetNumber.
     *
     * @param columnMapper   column mapping information
     * @param workbook       given workbook
     * @param sheetNumber    given sheet number
     * @param beginRowNumber row number to begin reading
     * @param readSize       row count to read
     * @return map list
     * @throws Exception exceptions
     */
    protected List<Map<String, Object>> readDataDirectly(ColumnMapper columnMapper, Workbook workbook, int sheetNumber,
                                                         int beginRowNumber, int readSize) throws Exception {
        if (columnMapper == null)
            throw new NullPointerException("columnMapper can not be null.");
        if (workbook == null)
            throw new NullPointerException("workbook can not be null.");
        if (sheetNumber < 0)
            throw new RuntimeException("sheet number can not below 0.");
        if (beginRowNumber < 0)
            throw new RuntimeException("beginRowNumber can not below 0.");
        if (readSize < 0)
            throw new RuntimeException("readSize can not below 0.");
        List<Map<String, Object>> mapList = new ArrayList<>();
        mapList.addAll(CustomImporter.getDataFromExcel(workbook, sheetNumber, columnMapper, beginRowNumber, readSize));
        return mapList;
    }

    /**
     * Read type data from a given workbook in sheet from a sheetNumber.
     *
     * @param clazz          type
     * @param columnMapper   column mapping information
     * @param workbook       given workbook
     * @param sheetNumber    given sheet number
     * @param beginRowNumber row number to begin reading
     * @param readSize       row count to read
     * @param <T>            type
     * @return type list
     * @throws Exception exceptions
     */
    @SuppressWarnings("unchecked")
    protected <T> List<T> readData(Class<T> clazz, ColumnMapper columnMapper, Workbook workbook, int sheetNumber,
                                   int beginRowNumber, int readSize)
            throws Exception {
        if (clazz == null || columnMapper == null || workbook == null || sheetNumber < 0 || beginRowNumber < 0
                || readSize <= 0)
            throw new RuntimeException("method parameter illegal.");
        return read(clazz, columnMapper, workbook, sheetNumber, beginRowNumber, readSize);
    }

    /**
     * Read type data from a given workbook in sheet from a sheetNumber.
     *
     * @param clazz          type
     * @param workbook       given workbook
     * @param sheetNumber    given sheet number
     * @param beginRowNumber row number to begin reading
     * @param readSize       row count to read
     * @param <T>            type
     * @return type list
     * @throws Exception exceptions
     */
    protected <T> List<T> readData(Class<T> clazz, Workbook workbook, int sheetNumber, int beginRowNumber, int readSize)
            throws Exception {
        if (clazz == null || workbook == null || sheetNumber < 0 || beginRowNumber < 0 || readSize <= 0)
            throw new RuntimeException("method parameter illegal.");
        ColumnMapper columnMapper = generateColumnMapper(clazz);
        return read(clazz, columnMapper, workbook, sheetNumber, beginRowNumber, readSize);
    }

    private <T> List<T> read(Class<T> clazz, ColumnMapper columnMapper, Workbook workbook, int sheetNumber,
                             int beginRowNumber, int readSize) throws Exception {
        List<T> dataList = new ArrayList<>();
        List<Map<Integer, Object>> mapList = new ArrayList<>();
        mapList.addAll(CustomImporter.getDataFromExcel(workbook, sheetNumber, beginRowNumber, readSize));
        for (Map<Integer, Object> dataMap : mapList) {
            Object object = ReflectionUtils.newObjectInstance(clazz);
            if (object == null)
                continue;
            List<Field> fieldList = ReflectionUtils.getFieldsUnStaticUnFinal(clazz);
            if (fieldList == null || fieldList.isEmpty())
                continue;
            for (Field field : fieldList) {
                String fieldName = field.getName();
                if (fieldName == null)
                    continue;
                Class<?> fieldJavaType = field.getType();
                if (fieldJavaType == null)
                    continue;
                Integer columnNumber = columnMapper.getColumnNumber(fieldName);
                if (columnNumber < 0)
                    continue;
                String fieldStrValue = String.valueOf(dataMap.get(columnNumber));
                if (fieldStrValue == null)
                    continue;
                Object fieldRealValue;
                try {
                    fieldRealValue = ValueUtils.parseValue(fieldStrValue,
                            ReflectionUtils.getClassByName(fieldJavaType.getName()));
                } catch (Exception e) {
                    fieldRealValue = null;
                }
                if (fieldRealValue == null)
                    continue;
                ReflectionUtils.setFieldValue(object, fieldName, fieldRealValue);
            }
            try {
                dataList.add((T) object);
            } catch (Exception e) {
            }
        }
        return dataList;
    }

    /**
     * Read data count from a given workbook in sheet from a sheetNumber where data begins at row.
     *
     * @param workbook           given workbook
     * @param sheetNumber        given sheet number
     * @param dataBeginRowNumber row number which data begins
     * @return data count
     * @throws Exception exceptions
     */
    protected int readDataCount(Workbook workbook, int sheetNumber, int dataBeginRowNumber) throws Exception {
        if (workbook == null)
            throw new NullPointerException("workbook can not be null.");
        if (sheetNumber < 0)
            throw new RuntimeException("sheet number can not below 0.");
        return CustomImporter.getDataCountFromExcel(workbook, sheetNumber, dataBeginRowNumber);
    }

    /**
     * Close workbook, release resources.
     *
     * @param workbook given workbook
     * @return true/false
     */
    protected boolean close(Workbook workbook) {
        return Common.closeWorkbook(workbook);
    }
}
