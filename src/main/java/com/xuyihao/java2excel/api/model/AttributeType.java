package com.xuyihao.java2excel.api.model;

import java.util.Map;

/**
 * Created by Xuyh at 2016/07/22 上午 10:42.
 *
 * 用来存储属性code, name, value的inner class
 */
public class AttributeType {
    /**
     * fields
     *
     * */
    private String attrId;
    private String attrCode;//属性编码
    private String attrName;//属性名
    private String attrType;//属性数据类型
    private String attrFormatRule;//属性取值范围及规则

    public AttributeType(){
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
}
