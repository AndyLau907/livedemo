package com.iceblue.livedemo.model.excel;

/**
 * @author Leiq
 * @program: LiveDemo
 * @description:
 * @date 2021-10-21 17:55:10
 */

public class ExcelWorkbookDesignerDataModel {
    private String name;
    private String capital;
    private String continent;
    private int area;
    private long population;

    public ExcelWorkbookDesignerDataModel(String name, String capital, String continent, int area, long population) {
        this.name = name;
        this.capital = capital;
        this.continent = continent;
        this.area = area;
        this.population = population;
    }

    public ExcelWorkbookDesignerDataModel() {
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCapital() {
        return capital;
    }

    public void setCapital(String capital) {
        this.capital = capital;
    }

    public String getContinent() {
        return continent;
    }

    public void setContinent(String continent) {
        this.continent = continent;
    }

    public int getArea() {
        return area;
    }

    public void setArea(int area) {
        this.area = area;
    }

    public long getPopulation() {
        return population;
    }

    public void setPopulation(long population) {
        this.population = population;
    }
}
