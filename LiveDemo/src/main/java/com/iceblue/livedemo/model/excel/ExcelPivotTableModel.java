package com.iceblue.livedemo.model.excel;

/**
 * 创建透视白的请求model
 */

public class ExcelPivotTableModel {
    private String excelVersion;
    private String averageName;
    private String sumName;
    private String maxName;
    private String minName;

    public String getExcelVersion() {
        return excelVersion;
    }

    public void setExcelVersion(String excelVersion) {
        this.excelVersion = excelVersion;
    }

    public String getAverageName() {
        return averageName;
    }

    public void setAverageName(String averageName) {
        this.averageName = averageName;
    }

    public String getSumName() {
        return sumName;
    }

    public void setSumName(String sumName) {
        this.sumName = sumName;
    }

    public String getMaxName() {
        return maxName;
    }

    public void setMaxName(String maxName) {
        this.maxName = maxName;
    }

    public String getMinName() {
        return minName;
    }

    public void setMinName(String minName) {
        this.minName = minName;
    }
}
