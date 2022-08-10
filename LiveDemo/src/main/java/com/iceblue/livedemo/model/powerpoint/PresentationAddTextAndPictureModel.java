package com.iceblue.livedemo.model.powerpoint;

import org.apache.commons.lang.StringUtils;

import java.awt.*;

/**
 * ppt添加文本或图片的请求model
 */

public class PresentationAddTextAndPictureModel {
    private String text;
    private String fontName;
    private String colorText;
    private float fontSize;
    private String addType;

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }


    public float getFontSize() {
        return fontSize;
    }

    public void setFontSize(float fontSize) {
        this.fontSize = fontSize;
    }

    public String getColorText() {
        return colorText;
    }

    public void setColorText(String colorText) {
        this.colorText = colorText;
    }

    public String getAddType() {
        return addType;
    }

    public void setAddType(String addType) {
        this.addType = addType;
    }
}
