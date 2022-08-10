package com.iceblue.livedemo.model.excel;

/**
 * @author Leiq
 * @program: LiveDemo
 * @description:
 * @date 2021-10-21 17:31:14
 */

public class ExcelPivotTableDataModel {
    private String name;
    private int vendorNo;
    private double sales;
    private int onHand;
    private int onOrder;
    private int area;
    private long population;

    public ExcelPivotTableDataModel(String name, int vendorNo, double sales, int onHand, int onOrder, int area, long population) {
        this.name = name;
        this.vendorNo = vendorNo;
        this.sales = sales;
        this.onHand = onHand;
        this.onOrder = onOrder;
        this.area = area;
        this.population = population;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getVendorNo() {
        return vendorNo;
    }

    public void setVendorNo(int vendorNo) {
        this.vendorNo = vendorNo;
    }

    public double getSales() {
        return sales;
    }

    public void setSales(double sales) {
        this.sales = sales;
    }

    public int getOnHand() {
        return onHand;
    }

    public void setOnHand(int onHand) {
        this.onHand = onHand;
    }

    public int getOnOrder() {
        return onOrder;
    }

    public void setOnOrder(int onOrder) {
        this.onOrder = onOrder;
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
