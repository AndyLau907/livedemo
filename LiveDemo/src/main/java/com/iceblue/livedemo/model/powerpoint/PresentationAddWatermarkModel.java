package com.iceblue.livedemo.model.powerpoint;

import com.iceblue.livedemo.utils.CommonHelper;
import org.apache.commons.lang.StringUtils;

/**
 * ppt添加水印的请求model
 */

public class PresentationAddWatermarkModel {
    private String text;
    private String fontName;
    private String colorText;
    private float fontSize;
    private float rotate;
    private String watermarkType;

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

    public String getColorText() {
        return colorText;
    }

    public void setColorText(String colorText) {
        this.colorText = colorText;
    }

    public String getWatermarkType() {
        return watermarkType;
    }

    public void setWatermarkType(String watermarkType) {
        this.watermarkType = watermarkType;
    }

    public float getFontSize() {
        return fontSize;
    }

    public void setFontSize(float fontSize) {
        this.fontSize = fontSize;
    }

    public float getRotate() {
        return rotate;
    }

    public void setRotate(float rotate) {
        this.rotate = rotate;
    }
}
