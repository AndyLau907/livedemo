package com.iceblue.livedemo.controller;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.powerpoint.*;
import com.iceblue.livedemo.service.PresentationDemoService;
import com.iceblue.livedemo.utils.FileInitHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.presentation.Presentation;
import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RequestMapping(value = "/ppt")
@RestController
public class PptController {

    @Autowired
    PresentationDemoService presentationDemoService;

    @RequestMapping(method = RequestMethod.POST, value = "/conversion")
    @ResponseBody
    public ResultModel conversion(@RequestPart(value = "pptFile", required = false) MultipartFile pdfFile, @RequestPart("jsonModel") PresentationConvertRequestModel model) {
        ResultModel resultModel = new ResultModel();

        if (StringUtils.isBlank(model.getConvertType())) {
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_SELECT_CONVERSION_TYPE);
            return resultModel;
        }
        Presentation ppt = FileInitHelper.initPPT(resultModel, pdfFile, "convert.pptx");
        if (ppt == null) {
            return resultModel;
        }
        try {
            presentationDemoService.preaentionConversion(resultModel, ppt, model.getConvertType());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/findOrReplace")
    @ResponseBody
    public ResultModel findOrReplace(@RequestPart(value = "pptFile") MultipartFile pptFile, @RequestPart("jsonModel") PresentationFindAndReplaceModel model) {
        ResultModel resultModel = new ResultModel();

        initFindAndReplaceModel(model);

        Presentation ppt = FileInitHelper.initPPT(resultModel, pptFile, "findAndReplace.pptx");
        if (ppt == null) {
            return resultModel;
        }
        try {
            presentationDemoService.findOrReplace(resultModel, ppt, model.getTypeChance(), model.getFindText(), model.getNewText());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/addTextOrImage")
    @ResponseBody
    public ResultModel addTextOrImage(@RequestPart(value = "pptImage", required = false) MultipartFile pptImage, @RequestPart("jsonModel") PresentationAddTextAndPictureModel model) {
        ResultModel resultModel = new ResultModel();

        initAddTextAndPictureModel(model);
        byte[] imgData = FileInitHelper.initImage(resultModel, pptImage, "Logo.png");
        try {
            presentationDemoService.addTextOrImage(resultModel, new Presentation(), model.getAddType(), model.getText(),
                    model.getFontName(), model.getColorText(), model.getFontSize(), imgData);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/addTextOrImageWatermark")
    @ResponseBody
    public ResultModel addTextOrImageWatermark(@RequestPart(value = "pptFile", required = false) MultipartFile pptFile,
                                               @RequestPart(value = "pptImage", required = false) MultipartFile pptImage,
                                               @RequestPart("jsonModel") PresentationAddWatermarkModel model) {
        ResultModel resultModel = new ResultModel();

        initAddWatermarkModel(model);
        Presentation ppt = FileInitHelper.initPPT(resultModel, pptFile, "findAndReplace.pptx");
        if (ppt == null) {
            return resultModel;
        }
        byte[] imgData = FileInitHelper.initImage(resultModel, pptImage, "Logo.png");
        try {
            presentationDemoService.addTextOrImageWatermark(resultModel, ppt, model, imgData);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/addCharts")
    @ResponseBody
    public ResultModel addCharts(@RequestPart("jsonModel") PresentationChartModel model) {
        ResultModel resultModel = new ResultModel();

        if (StringUtils.isBlank(model.getChartName())) {
            model.setChartName("ColumnClustered");
        }
        try {
            presentationDemoService.charts(resultModel, model.getChartName());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getConvertTypeOptions")
    public ResultModel getConvertTypeOptions() {
        ResultModel resultModel = new ResultModel();
        presentationDemoService.getConvertTypeOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getPptChartOptions")
    public ResultModel getPptChartOptions() {
        ResultModel resultModel = new ResultModel();
        presentationDemoService.getPptChartOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getRotateOptions")
    public ResultModel getRotateOptions() {
        ResultModel resultModel = new ResultModel();
        presentationDemoService.getRotateOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getChartData")
    public ResultModel getChartData() {
        ResultModel resultModel = new ResultModel();
        presentationDemoService.getChartData(resultModel);
        return resultModel;
    }

    /*******************************************************************************************************/
    private void initFindAndReplaceModel(PresentationFindAndReplaceModel model){
        if (StringUtils.isBlank(model.getTypeChance())) {
            model.setTypeChance("FIND");
        }
        if (StringUtils.isBlank(model.getFindText())) {
            model.setFindText("");
        }
        if (StringUtils.isBlank(model.getNewText())) {
            model.setNewText("");
        }
    }
    private void initAddWatermarkModel(PresentationAddWatermarkModel model) {
        if (model.getFontSize() == 0) {
            model.setFontSize(20f);
        }
        if (StringUtils.isBlank(model.getWatermarkType())) {
            model.setWatermarkType("TEXT");
        }
        if (StringUtils.isBlank(model.getText())) {
            model.setText("E-iceblue");
        }
        if (StringUtils.isBlank(model.getColorText())) {
            model.setColorText("000000");
        }
        if (StringUtils.isBlank(model.getFontName())) {
            model.setFontName("Arial");
        }
    }

    private void initAddTextAndPictureModel(PresentationAddTextAndPictureModel model) {
        if (model.getFontSize() == 0) {
            model.setFontSize(20f);
        }
        if (StringUtils.isBlank(model.getAddType())) {
            model.setAddType("TEXT");
        }
        if (StringUtils.isBlank(model.getText())) {
            model.setText("E-iceblue");
        }
        if (StringUtils.isBlank(model.getColorText())) {
            model.setColorText("000000");
        }
        if (StringUtils.isBlank(model.getFontName())) {
            model.setFontName("Arial");
        }
    }
}
