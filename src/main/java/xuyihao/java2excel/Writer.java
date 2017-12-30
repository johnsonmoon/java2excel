package xuyihao.java2excel;

import xuyihao.java2excel.core.entity.model.Attribute;
import xuyihao.java2excel.core.entity.model.Template;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Created by Xuyh at 2017/12/30 下午 03:15.
 */
public class Writer<T> {
    private Template template;

    public boolean writeExcel() {


        return true;
    }

    public boolean writeExcelWithData(List<T> tList) {


        return true;
    }

    /**
     * Using java reflect and annotation to generate template information
     *
     * @param tClass
     */
    private void generateTemplateLess(Class<T> tClass) {
        //TODO 定义注解来定义业务模型和template，模型字段和attribute之间的映射关系
        this.template = new Template();
        template.setName(tClass.getName());
        Field[] fields = tClass.getDeclaredFields();
        for (Field field : fields) {
            String fieldStr = field.toString();
            if (fieldStr.contains("static")
                    || fieldStr.contains("final"))
                continue;
            Attribute attribute = new Attribute();
            attribute.setAttrCode(field.getName());
            attribute.setAttrName(field.getName());
            String typeStr = field.getType().toString();
            attribute.setAttrType(typeStr.substring(typeStr.lastIndexOf(".")+1, typeStr.length()));
        }
    }
}
