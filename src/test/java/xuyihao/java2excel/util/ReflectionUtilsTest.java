package xuyihao.java2excel.util;

import org.junit.Test;
import xuyihao.java2excel.Reader;
import xuyihao.java2excel.entity.BaseTypes;
import xuyihao.java2excel.entity.Student;
import xuyihao.java2excel.entity.Teacher;

import java.lang.reflect.Field;

/**
 * Created by xuyh at 2018/1/3 11:34.
 */
public class ReflectionUtilsTest {
    @Test
    public void testClassNames() {
        Class<?> clazz = Teacher.class;

        System.out.println(ReflectionUtils.getClassNameShort(clazz));
        System.out.println(ReflectionUtils.getClassNameEntire(clazz));

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            System.out.println(ReflectionUtils.getFieldTypeNameShort(field));
            System.out.println(ReflectionUtils.getFieldTypeNameEntire(field));
        }
    }

    @Test
    public void testSetFieldValue() {
        Student student = new Student();
        student.setName("Johnson");

        System.out.println(ReflectionUtils.getFieldValue(student, "name"));
        System.out.println(ReflectionUtils.setFieldValue(student, "name", "John"));
        System.out.println(student.getName());
    }

    @Test
    public void testGetClassType() throws Exception {
        String className1 = String.class.getName();
        String className2 = "String";
        String className3 = "java.lang.Float";

        try {
            Class<?> clazz = ClassLoader.getSystemClassLoader().loadClass(className1);
            System.out.println(clazz.getName());
        } catch (Exception e) {
            System.out.println("class not found");
        }

        try {
            Class<?> clazz = Class.forName(className2);
            System.out.println(clazz.getName());
        } catch (Exception e) {
            System.out.println("class not found");
        }

        try {
            Class<?> clazz = Class.forName(className3);
            System.out.println(clazz.getName());
        } catch (Exception e) {
            System.out.println("class not found");
        }
    }

    @Test
    public void testGetBaseType() throws Exception {
        String className = "double";
        String value = "12.36";

        try {
            Class<?> clazz = ReflectionUtils.getClassByName(className);
            double var = (double) Reader.formatValue(value, clazz);
            System.out.println(var);
        } catch (Exception e) {
            System.out.println("class not found");
        }
    }

    @Test
    public void testGetBaseTypeNames() throws Exception {
        //Integer, Float, Byte, Double, Boolean, Character, Long, Short
        for (Field field : ReflectionUtils.getFieldsAll(BaseTypes.class)) {
            String nameEntire = ReflectionUtils.getFieldTypeNameEntire(field);
            System.out.println(nameEntire);

            Class<?> tClass = ReflectionUtils.getClassByName(nameEntire);
            System.out.println(tClass.getName());
        }
    }
}
