package com.xuyihao.java2excel.api.model.outer;

import org.springframework.data.mongodb.core.index.CompoundIndex;
import org.springframework.data.mongodb.core.index.CompoundIndexes;

import java.util.Map;

/**
 * 类的属性定义
 *
 * @author luhy Create at 2016年4月19日 下午5:21:25
 */
@CompoundIndexes({ @CompoundIndex(name = "idx_t_c_c", def = "{tenant:1, classCode:1, code:1}") })
public class Attribute implements TenantEnable{// Dao处理数据保存，查询时，可根据自接口自动增加租户信息
    public static final String ATTR_NAME = "Y_name";
    public static final String ATTR_STATE = "Y_state";

    private String id;
    private String tenant;
    private String classCode;
    private String code;
    private String name;
    private String type;
    private String defaultValue;
    private boolean required;
    private boolean builtin;
    private String unit;

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit;
    }

    /**
     * <pre>
     *params:   属性配参数
     *params.values 属性取值列表(单选，多选，列表选择（单选）)
     *params.max 最大值（数字，日期，时间）
     *params.min 最小值（数字，日期，时间）
     *params.maxLength  文本最大长度
     *params.minLength  文本最小长度
     *params.regexType  文本正则类型（前端较验，台端不较验，前端应事先定义好类型。比如（name:ip较验，regex:xxx）
     *prams.regexError  文本正则不匹配时提示消息。
     * </pre>
     */
    private Map<String, Object> params;
    private int sortIndex;
    private String descr;

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public boolean isBuiltin() {
        return builtin;
    }

    public void setBuiltin(boolean builtin) {
        this.builtin = builtin;
    }

    public String getTenant() {
        return tenant;
    }

    public void setTenant(String tenant) {
        this.tenant = tenant;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public Map<String, Object> getParams() {
        return params;
    }

    public void setParams(Map<String, Object> params) {
        this.params = params;
    }

    public int getSortIndex() {
        return sortIndex;
    }

    public void setSortIndex(int sortIndex) {
        this.sortIndex = sortIndex;
    }

    public String getDescr() {
        return descr;
    }

    public void setDescr(String descr) {
        this.descr = descr;
    }

    public String getClassCode() {
        return classCode;
    }

    public void setClassCode(String classCode) {
        this.classCode = classCode;
    }

    public boolean equals(Object obj) {
        if (obj == null || !(obj instanceof Attribute))
            return false;
        Attribute other = (Attribute) obj;
        return this.toString().equals(other.toString());
    }

    /**
     * <pre>
     * 筒单返回配置项类型Json格式数据,用于调试，查看<br>
     * <font color=red>
     * 实际生产环境需要Json数据时，当用Json工且格式化此对象，以获取真实Json
     * </font>
     * </pre>
     */
    public String toString() {
        StringBuilder sb = new StringBuilder("{");
        sb.append("\"").append("id").append("\":").append(" \"").append(id).append("\", ");
        sb.append("\"").append("tenant").append("\":").append(" \"").append(tenant).append("\", ");
        sb.append("\"").append("classCode").append("\":").append(" \"").append(classCode).append("\", ");
        sb.append("\"").append("code").append("\":").append(" \"").append(code).append("\", ");
        sb.append("\"").append("name").append("\":").append(" \"").append(name).append("\", ");
        sb.append("\"").append("type").append("\":").append(" \"").append(type).append("\", ");
        sb.append("\"").append("descr").append("\":").append(" \"").append(descr).append("\", ");
        sb.append("\"").append("defaultValue").append("\":").append(" \"").append(defaultValue).append("\", ");
        sb.append("\"").append("required").append("\":").append(" ").append(required).append(", ");
        sb.append("\"").append("builtin").append("\":").append(" ").append(builtin).append(", ");
        sb.append("\"").append("unit").append("\":").append(" \"").append(unit).append("\", ");
        sb.append("\"").append("sortIndex").append("\":").append(" ").append(sortIndex).append(", ");
        sb.append("\"").append("params").append("\":").append(" \"").append(params == null ? null : params.toString())
                .append("\", ");
        sb.append("}");

        return sb.toString();
    }
}
