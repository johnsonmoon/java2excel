package com.xuyihao.java2excel.api.model;

import com.xuyihao.java2excel.api.model.outer.CIFormContainer;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/21 下午 01:40.
 * 用来导出导入Ci数据到Excel文件中的数据模型
 * 初期会跟CIForm相同，后期会有所改动
 */
public class ExcelCiTemplate {
    /**
     * fields
     *
     * */
    private String id;//id 这个ID就是ciForm的ID
    private String tenant;//tenant 租户ID
    private String classCode;//classCode 类型编码(Template的code)
    private List<CIFormContainer> attrsForm;//attrsForm 保存属性的List

    /**
     * getters & setters
     *
     * */
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

    public List<CIFormContainer> getAttrsForm(){
        return this.attrsForm == null ? null : Collections.unmodifiableList(this.attrsForm);
    }

    public void setAttrsForm(List<CIFormContainer> attrsForm){
        if(this.attrsForm != null){
            this.attrsForm = new ArrayList<CIFormContainer>(attrsForm);
        }
    }
}
