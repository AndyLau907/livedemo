package com.iceblue.livedemo.model.excel;

/**
 * @author Leiq
 * @program: LiveDemo
 * @description:
 * @date 2021-10-19 11:52:41
 */

public class ExcelCalculateFormulasModel {
    private String excelVersion;
    private String mathFunction;
    private float mathData;
    private String logicFunction;
    private boolean logicValueOne;
    private boolean logicValueTwo;
    private String simpleExpression;
    private int fData;
    private int sData;
    private String midText;
    private int startNumber;
    private int numberChart;

    public String getExcelVersion() {
        return excelVersion;
    }

    public void setExcelVersion(String excelVersion) {
        this.excelVersion = excelVersion;
    }

    public String getMathFunction() {
        return mathFunction;
    }

    public void setMathFunction(String mathFunction) {
        this.mathFunction = mathFunction;
    }

    public float getMathData() {
        return mathData;
    }

    public void setMathData(float mathData) {
        this.mathData = mathData;
    }

    public String getLogicFunction() {
        return logicFunction;
    }

    public void setLogicFunction(String logicFunction) {
        this.logicFunction = logicFunction;
    }

    public boolean isLogicValueOne() {
        return logicValueOne;
    }

    public void setLogicValueOne(boolean logicValueOne) {
        this.logicValueOne = logicValueOne;
    }

    public boolean isLogicValueTwo() {
        return logicValueTwo;
    }

    public void setLogicValueTwo(boolean logicValueTwo) {
        this.logicValueTwo = logicValueTwo;
    }

    public String getSimpleExpression() {
        return simpleExpression;
    }

    public void setSimpleExpression(String simpleExpression) {
        this.simpleExpression = simpleExpression;
    }

    public int getfData() {
        return fData;
    }

    public void setfData(int fData) {
        this.fData = fData;
    }

    public int getsData() {
        return sData;
    }

    public void setsData(int sData) {
        this.sData = sData;
    }

    public String getMidText() {
        return midText;
    }

    public void setMidText(String midText) {
        this.midText = midText;
    }

    public int getStartNumber() {
        return startNumber;
    }

    public void setStartNumber(int startNumber) {
        this.startNumber = startNumber;
    }

    public int getNumberChart() {
        return numberChart;
    }

    public void setNumberChart(int numberChart) {
        this.numberChart = numberChart;
    }
}
