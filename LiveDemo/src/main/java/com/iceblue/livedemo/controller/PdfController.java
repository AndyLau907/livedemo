package com.iceblue.livedemo.controller;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.pdf.PdfAddWaterMarkRequestModel;
import com.iceblue.livedemo.model.pdf.PdfConvertRequestModel;
import com.iceblue.livedemo.model.pdf.PdfFindAndHighlightRequestModel;
import com.iceblue.livedemo.model.pdf.PdfTableRequestModel;
import com.iceblue.livedemo.service.PdfDemoService;
import com.iceblue.livedemo.utils.CommonHelper;
import com.iceblue.livedemo.utils.FileInitHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.pdf.PdfDocument;
import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RequestMapping(value = "/pdf")
@RestController
public class PdfController {

    @Autowired
    PdfDemoService pdfDemoService;

    @RequestMapping(method = RequestMethod.POST, value = "/conversion")
    @ResponseBody
    public ResultModel conversion(@RequestPart(value = "pdfFile", required = false) MultipartFile pdfFile, @RequestPart("jsonModel") PdfConvertRequestModel model) {
        ResultModel resultModel = new ResultModel();

        if (StringUtils.isBlank(model.getConvertType())) {
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_SELECT_CONVERSION_TYPE);
            return resultModel;
        }
        PdfDocument pdf = FileInitHelper.initPDF(resultModel, pdfFile, "conversion.pdf");
        if (pdf == null) {
            return resultModel;
        }
        try {
            pdfDemoService.pdfConversion(resultModel, pdf, pdfFile, model.getConvertType());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/drawTable")
    @ResponseBody
    public ResultModel drawTable(@RequestPart("jsonModel") PdfTableRequestModel model) {
        ResultModel resultModel = new ResultModel();

        try {
            pdfDemoService.drawTable(resultModel, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(value = "/mergePdf")
    @ResponseBody
    public ResultModel mergePdf(@RequestParam(value = "pdfFiles", required = true) MultipartFile[] pdfFile) {
        ResultModel resultModel = new ResultModel();

        if (pdfFile == null || pdfFile.length == 0) {
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        try {
            pdfDemoService.mergePdf(resultModel, pdfFile);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/addWaterMark")
    @ResponseBody
    public ResultModel addWaterMark(@RequestPart(value = "pdfFile", required = false) MultipartFile pdfFile,
                                    @RequestPart(value = "waterMarkImg", required = false) MultipartFile waterMarkImg,
                                    @RequestPart("jsonModel") PdfAddWaterMarkRequestModel model) {
        ResultModel resultModel = new ResultModel();

        initWaterMarkModel(model);
        PdfDocument pdf = FileInitHelper.initPDF(resultModel, pdfFile, "waterMark.pdf");
        if (pdf == null) {
            return resultModel;
        }
        byte[] imgData = FileInitHelper.initImage(resultModel, waterMarkImg, "E-logo.png");
        if (imgData == null) {
            return resultModel;
        }
        try {
            pdfDemoService.addWaterMark(resultModel, pdf, imgData, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/findAndHighlight")
    @ResponseBody
    public ResultModel findAndHighlight(@RequestPart(value = "pdfFile", required = false) MultipartFile pdfFile, @RequestPart("jsonModel") PdfFindAndHighlightRequestModel model) {
        ResultModel resultModel = new ResultModel();

        initFindAndHighlightModel(model);
        PdfDocument pdf = FileInitHelper.initPDF(resultModel, pdfFile, "findAndHighlight.pdf");
        if (pdf == null) {
            return resultModel;
        }
        try {
            pdfDemoService.findAndHighlight(resultModel, pdf, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getConvertOptions")
    public ResultModel getConvertOptions() {
        ResultModel resultModel = new ResultModel();
        pdfDemoService.getConvertOptions(resultModel);
        return resultModel;
    }
    @RequestMapping(method = RequestMethod.GET, value = "/getTableData")
    public ResultModel getTableData() {
        ResultModel resultModel = new ResultModel();
        pdfDemoService.getTableData(resultModel);
        return resultModel;
    }
    /********************************************************************************************/
    private void initWaterMarkModel(PdfAddWaterMarkRequestModel model) {
        if (model.getTextSize() == 0) {
            model.setTextSize(20);
        }
        if (StringUtils.isBlank(model.getTextFont())) {
            model.setTextFont("Arial");
        }
        if (StringUtils.isBlank(model.getWaterMarkText())) {
            model.setWaterMarkText("E-iceblue");
        }
        //如果颜色值为空 或者解析颜色错误 那么默认字体颜色为黑色
        if (StringUtils.isBlank(model.getTextColor()) || CommonHelper.String2Color(model.getTextColor()) == null) {
            model.setTextColor("000000");
        }
        if (StringUtils.isBlank(model.getWaterMarkType())) {
            model.setWaterMarkType("Text");
        }
    }

    private void initFindAndHighlightModel(PdfFindAndHighlightRequestModel model) {
        if (StringUtils.isBlank(model.getText())) {
            model.setText("");
        }
        if (StringUtils.isBlank(model.getColor()) || CommonHelper.String2Color(model.getColor()) == null) {
            //默认黄色
            model.setColor("ffff00");
        }
    }
}
