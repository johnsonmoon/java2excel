package xuyihao.java2excel.core.entity.model;

/**
 * 对象属性，对应模型中的字段，excel表头
 *
 * Created by Xuyh at 2016/07/22 上午 10:42.
 */
public class Attribute {
    private String attrCode;//属性码
    private String attrName;//属性名
    private String attrType;//属性类型
    private String defaultValue;//属性默认值
    private String unit;//属性单位

    public Attribute() {
    }

    public Attribute(String attrCode, String attrName, String attrType, String defaultValue, String unit) {
        this.attrCode = attrCode;
        this.attrName = attrName;
        this.attrType = attrType;
        this.defaultValue = defaultValue;
        this.unit = unit;
    }

    public String getAttrCode() {
        return attrCode;
    }

    public void setAttrCode(String attrCode) {
        this.attrCode = attrCode;
    }

    public String getAttrName() {
        return attrName;
    }

    public void setAttrName(String attrName) {
        this.attrName = attrName;
    }

    public String getAttrType() {
        return attrType;
    }

    public void setAttrType(String attrType) {
        this.attrType = attrType;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit;
    }
}
