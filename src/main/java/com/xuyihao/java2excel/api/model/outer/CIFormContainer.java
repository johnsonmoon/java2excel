package com.xuyihao.java2excel.api.model.outer;

/**
 * Created by Xuyh at 2016/07/21 下午 06:08.
 */
public class CIFormContainer {
    public static final String TYPE_ATTR = "attribute";
    public static final String TYPE_GROUPLINE = "groupLine";
    private String name;
    private String code;
    private String type;
    private String id;

    public CIFormContainer() {
    }

    public static CIFormContainer attributeContainer(String id, String code) {
        return new CIFormContainer(null, TYPE_ATTR, code, null);
    }

    public CIFormContainer(String id, String type, String code, String name) {
        if (id == null && code != null) {
            this.id = code;
        } else {
            this.id = id;
        }
        this.type = type;
        this.code = code;
        this.name = name;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String toString() {
        return String.format("{\"type\":\"%s\", \"code\":\"%s\", \"name\":\"%s\",}", type, code, name);
    }
}
