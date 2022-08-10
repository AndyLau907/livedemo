package com.iceblue.livedemo.model.powerpoint;

/**
 * ppt转换请求model
 */

public class PresentationConvertRequestModel {
    /**
     * 需要转换的类型
     */
    private String convertType;

    public String getConvertType() {
        return convertType;
    }

    public void setConvertType(String convertType) {
        this.convertType = convertType;
    }
}
