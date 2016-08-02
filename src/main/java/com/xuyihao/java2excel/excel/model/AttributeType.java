package com.xuyihao.java2excel.excel.model;

/**
 * Created by Xuyh at 2016/07/22 上午 10:42.
 *
 * 用来存储属性code, name, default value的class
 */
public class AttributeType {
    /**
     * fields
     *
     * */
    private String attrId;
    private String attrCode;
    private String attrName;
    private String attrType;
    private String attrFormatRule;
    private String defaultValue;
    private String unit;

    public AttributeType(){
        this.attrId = "";
        this.attrCode = "";
        this.attrName = "";
        this.attrType = "";
        this.attrFormatRule = "";
        this.defaultValue = "";
        this.unit = "";
    }

    public AttributeType(AttributeType attributeType){
        this.attrId = attributeType.getAttrId();
        this.attrCode = attributeType.getAttrCode();
        this.attrName = attributeType.getAttrName();
        this.attrType = attributeType.getAttrType();
        this.attrFormatRule = attributeType.getAttrFormatRule();
        this.defaultValue = attributeType.getDefaultValue();
        this.unit = attributeType.getUnit();
    }

    /**
     * getters & setters
     *
     */
    public void setAttrId(String attrId){
        this.attrId = attrId;
    }

    public String getAttrId(){
        return this.attrId;
    }

    public void setAttrName(String attrName){
        this.attrName = attrName;
    }

    public String getAttrName(){
        return this.attrName;
    }

    public void setAttrCode(String attrCode){
        this.attrCode = attrCode;
    }

    public String getAttrCode(){
        return this.attrCode;
    }

    public void setAttrType(String attrType){
        this.attrType = attrType;
    }

    public String getAttrType(){
        return this.attrType;
    }

    public void setAttrFormatRule(String formatRule){
        this.attrFormatRule = formatRule;
    }

    public String getAttrFormatRule(){
        return this.attrFormatRule;
    }

    public void setDefaultValue(String defaultValue){
        this.defaultValue = defaultValue;
    }

    public String getDefaultValue(){
        return this.defaultValue;
    }

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit;
    }
}
