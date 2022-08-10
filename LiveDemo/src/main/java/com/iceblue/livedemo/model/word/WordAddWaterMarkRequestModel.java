package com.iceblue.livedemo.model.word;

import org.apache.commons.lang.StringUtils;

public class WordAddWaterMarkRequestModel {
    private String waterMarkType;
    private String waterMarkText;
    private String textColor;
    private String textFont;
    private int textSize;

    public String getTextFont() {
        return textFont;
    }

    public void setTextFont(String textFont) {
        this.textFont = textFont;
    }

    public String getWaterMarkType() {
        return waterMarkType;
    }

    public void setWaterMarkType(String waterMarkType) {
        this.waterMarkType = waterMarkType;
    }

    public String getWaterMarkText() {
        return waterMarkText;
    }

    public void setWaterMarkText(String waterMarkText) {
        this.waterMarkText = waterMarkText;
    }

    public String getTextColor() {
        return textColor;
    }

    public void setTextColor(String textColor) {
        this.textColor = textColor;
    }

    public int getTextSize() {
        return textSize;
    }

    public void setTextSize(int textSize) {
        this.textSize = textSize;
    }
}
