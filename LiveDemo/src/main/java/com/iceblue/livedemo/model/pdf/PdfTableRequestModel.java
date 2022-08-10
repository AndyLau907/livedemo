package com.iceblue.livedemo.model.pdf;

import org.apache.commons.lang.StringUtils;

public class PdfTableRequestModel {
    private String borderColor;
    private boolean repeatHeader;

    public String getBorderColor() {
        return borderColor;
    }

    public void setBorderColor(String borderColor) {
        this.borderColor = borderColor;
    }

    public boolean isRepeatHeader() {
        return repeatHeader;
    }

    public void setRepeatHeader(boolean repeatHeader) {
        this.repeatHeader = repeatHeader;
    }
}
