package xuyihao.java2excel.excel.entity;

/**
 * 用来存储属性code, name, default value的class。
 *
 * Created by Xuyh at 2016/07/22 上午 10:42.
 */
public class AttributeType {
    public final static String NOT_NULL = "非空";
    private String attrId;
    private String attrCode; // 属性编码
    private String attrName; // 属性名
    private String attrType; // 属性数据类型
    private boolean required; // 是否非空
    private String attrFormatRule; // 属性取值范围及规则
    private String defaultValue; // 该属性默认值
    private String unit; // 该属性的计量单位

    public AttributeType() {
        this.attrId = "";
        this.attrCode = "";
        this.attrName = "";
        this.attrType = "";
        this.required = false;
        this.attrFormatRule = "";
        this.defaultValue = "";
        this.unit = "";
    }

    public AttributeType(AttributeType attributeType) {
        this.attrId = attributeType.getAttrId();
        this.attrCode = attributeType.getAttrCode();
        this.attrName = attributeType.getAttrName();
        this.attrType = attributeType.getAttrType();
        this.required = attributeType.isRequired();
        this.attrFormatRule = attributeType.getAttrFormatRule();
        this.defaultValue = attributeType.getDefaultValue();
        this.unit = attributeType.getUnit();
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public void setAttrId(String attrId) {
        this.attrId = attrId;
    }

    public String getAttrId() {
        return this.attrId;
    }

    public void setAttrName(String attrName) {
        this.attrName = attrName;
    }

    public String getAttrName() {
        return this.attrName;
    }

    public void setAttrCode(String attrCode) {
        this.attrCode = attrCode;
    }

    public String getAttrCode() {
        return this.attrCode;
    }

    public void setAttrType(String attrType) {
        this.attrType = attrType;
    }

    public String getAttrType() {
        return this.attrType;
    }

    public void setAttrFormatRule(String formatRule) {
        this.attrFormatRule = formatRule;
    }

    public String getAttrFormatRule() {
        return this.attrFormatRule;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public String getDefaultValue() {
        return this.defaultValue;
    }

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit;
    }
}