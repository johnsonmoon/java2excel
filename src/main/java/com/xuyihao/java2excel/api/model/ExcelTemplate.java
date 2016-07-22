package com.xuyihao.java2excel.api.model;

import org.w3c.dom.Attr;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 下午 01:40.
 *
 * 用来导出导入数据到Excel文件中的数据模型
 */
public class ExcelTemplate {
    /**
     * fields
     *
     * */
    private String id;//id 这个ID就是ciForm的ID
    private String tenant;
    private String classCode;//类型编码(Template的code)
    private String className;//类型的名称（Template的name）
    private List<AttributeType> attrbuteTypes;
    private List<String> attrValues;

    /**
     * getters & setters
     *
     * */
    public void setAttrbuteTypes(List<AttributeType> attrbuteTypes){
        this.attrbuteTypes = attrbuteTypes;
    }

    public List<AttributeType> getAttrbuteTypes(){
        return this.attrbuteTypes;
    }

    public void

    public String getId(){
        return this.id;
    }

    public void setId(String id){
        this.id = id;
    }

    public String getTenant(){
        return this.tenant;
    }

    public void setTenant(String tenant){
        this.tenant = tenant;
    }

    public String getClassCode(){
        return this.classCode;
    }

    public void setClassCode(String classCode){
        this.classCode = classCode;
    }

    public String getClassName(){
        return this.className;
    }

    public void setClassName(String name){
        this.className = name;
    }

    public AttributeType
}
