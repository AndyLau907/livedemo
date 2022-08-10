package com.iceblue.livedemo.model.word;

import org.apache.commons.lang.StringUtils;

/**
 *
 */
public class WordHeaderAndFooterRequestModel {
    private String pageSize;
    //Header
    private String headerText;
    private String headerFont;
    private int headerFontSize;
    private String headerTextColor;
    private String headerTextAlignment;
    //Footer
    private String footerText;
    private String footerFont;
    private int footerFontSize;
    private String footerTextColor;
    private String footerTextAlignment;

    public String getPageSize() {
        return pageSize;
    }

    public void setPageSize(String pageSize) {
        this.pageSize = pageSize;
    }

    public String getHeaderText() {
        return headerText;
    }

    public void setHeaderText(String headerText) {
        this.headerText = headerText;
    }

    public String getHeaderFont() {
        return headerFont;
    }

    public void setHeaderFont(String headerFont) {
        this.headerFont = headerFont;
    }

    public int getHeaderFontSize() {
        return headerFontSize;
    }

    public void setHeaderFontSize(int headerFontSize) {
        this.headerFontSize = headerFontSize;
    }

    public String getHeaderTextColor() {
        return headerTextColor;
    }

    public void setHeaderTextColor(String headerTextColor) {
        this.headerTextColor = headerTextColor;
    }

    public String getHeaderTextAlignment() {
        return headerTextAlignment;
    }

    public void setHeaderTextAlignment(String headerTextAlignment) {
        this.headerTextAlignment = headerTextAlignment;
    }

    public String getFooterText() {
        return footerText;
    }

    public void setFooterText(String footerText) {
        this.footerText = footerText;
    }

    public String getFooterFont() {
        return footerFont;
    }

    public void setFooterFont(String footerFont) {
        this.footerFont = footerFont;
    }

    public int getFooterFontSize() {
        return footerFontSize;
    }

    public void setFooterFontSize(int footerFontSize) {
        this.footerFontSize = footerFontSize;
    }

    public String getFooterTextColor() {
        return footerTextColor;
    }

    public void setFooterTextColor(String footerTextColor) {
        this.footerTextColor = footerTextColor;
    }

    public String getFooterTextAlignment() {
        return footerTextAlignment;
    }

    public void setFooterTextAlignment(String footerTextAlignment) {
        this.footerTextAlignment = footerTextAlignment;
    }
}
