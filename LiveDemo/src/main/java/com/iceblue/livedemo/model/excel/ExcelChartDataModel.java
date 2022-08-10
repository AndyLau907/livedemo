package com.iceblue.livedemo.model.excel;

/**
 * @author Leiq
 * @program: LiveDemo
 * @description:
 * @date 2021-10-21 17:13:30
 */

public class ExcelChartDataModel {
    private String country;
    private double jun;
    private double aug;

    public ExcelChartDataModel(String country, double jun, double aug) {
        this.country = country;
        this.jun = jun;
        this.aug = aug;
    }

    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public double getJun() {
        return jun;
    }

    public void setJun(double jun) {
        this.jun = jun;
    }

    public double getAug() {
        return aug;
    }

    public void setAug(double aug) {
        this.aug = aug;
    }
}
