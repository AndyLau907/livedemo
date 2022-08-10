package com.iceblue.livedemo.model.powerpoint;

/**
 * @author Leiq
 * @program: LiveDemo
 * @description:
 * @date 2021-10-21 16:01:56
 */

public class ReportModel {
    private String salesPers;
    private int saleAmt;
    private int comPct;
    private int comAmt;

    public String getSalesPers() {
        return salesPers;
    }

    public void setSalesPers(String salesPers) {
        this.salesPers = salesPers;
    }

    public int getSaleAmt() {
        return saleAmt;
    }

    public void setSaleAmt(int saleAmt) {
        this.saleAmt = saleAmt;
    }

    public int getComPct() {
        return comPct;
    }

    public void setComPct(int comPct) {
        this.comPct = comPct;
    }

    public int getComAmt() {
        return comAmt;
    }

    public void setComAmt(int comAmt) {
        this.comAmt = comAmt;
    }
}
