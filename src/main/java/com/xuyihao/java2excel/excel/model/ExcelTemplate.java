package com.xuyihao.java2excel.excel.model;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 下午 01:40.
 *
 * 用来导出导入数据到Excel文件中的数据模型
 */
public class ExcelTemplate {

    private String id;
    private String tenant;
    private String classCode;
    private String className;
    private List<AttributeType> attrbuteTypes;
    private List<String> attrValues;

    public ExcelTemplate() {
        this.id = "";
        this.tenant = "";
        this.classCode = "";
        this.className = "";
        this.attrbuteTypes = new ArrayList<AttributeType>();
        this.attrValues = new ArrayList<String>();
    }

    public ExcelTemplate(ExcelTemplate template) {
        this.id = template.getId();
        this.tenant = template.getTenant();
        this.classCode = template.getClassCode();
        this.className = template.getClassName();
        this.attrbuteTypes = template.getAttrbuteTypes();
        this.attrValues = template.getAttrValues();
    }

    public void setAttrbuteTypes(List<AttributeType> attrbuteTypes) {
        this.attrbuteTypes = attrbuteTypes;
    }

    public List<AttributeType> getAttrbuteTypes() {
        return this.attrbuteTypes;
    }

    public void setAttrValues(List<String> attrValues) {
        this.attrValues = attrValues;
    }

    public List<String> getAttrValues() {
        return this.attrValues;
    }

    public String getId() {
        return this.id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getTenant() {
        return this.tenant;
    }

    public void setTenant(String tenant) {
        this.tenant = tenant;
    }

    public String getClassCode() {
        return this.classCode;
    }

    public void setClassCode(String classCode) {
        this.classCode = classCode;
    }

    public String getClassName() {
        return this.className;
    }

    public void setClassName(String name) {
        this.className = name;
    }
}
