package com.iceblue.livedemo.model;

/**
 * 用于返回各种下拉列表选项的模型
 */
public class TypeOptionsModel {
    private String typeViewName;
    private String typeValue;

    public TypeOptionsModel(){

    }
    public TypeOptionsModel(String typeViewName, String typeValue) {
        this.typeViewName = typeViewName;
        this.typeValue = typeValue;
    }

    public String getTypeViewName() {
        return typeViewName;
    }

    public void setTypeViewName(String typeViewName) {
        this.typeViewName = typeViewName;
    }

    public String getTypeValue() {
        return typeValue;
    }

    public void setTypeValue(String typeValue) {
        this.typeValue = typeValue;
    }
}
