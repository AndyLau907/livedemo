package com.iceblue.livedemo.controller;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.word.*;
import com.iceblue.livedemo.service.WordDemoService;
import com.iceblue.livedemo.utils.CommonHelper;
import com.iceblue.livedemo.utils.FileInitHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.doc.Document;
import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@RequestMapping(value = "/word")
@RestController
public class WordController {
    @Autowired
    WordDemoService wordDemoService;

    @RequestMapping(method = RequestMethod.POST, value = "/conversion")
    @ResponseBody
    public ResultModel conversion(@RequestPart(value = "docFile", required = false) MultipartFile docFile, @RequestPart("jsonModel") WordConvertRequestModel model) {
        ResultModel resultModel = new ResultModel();

        if (StringUtils.isBlank(model.getConvertType())) {
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_SELECT_CONVERSION_TYPE);
            return resultModel;
        }
        Document doc = FileInitHelper.initWord(resultModel, docFile, "conversion.doc");
        if (doc == null) {
            return resultModel;
        }
        try {
            wordDemoService.wordConversion(resultModel, doc, model.getConvertType());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        } finally {
            doc.close();
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/findAndHighlight")
    @ResponseBody
    public ResultModel findAndHighlight(@RequestPart(value = "docFile", required = false) MultipartFile docFile, @RequestPart("jsonModel") WordFindAndHighlightRequestModel model) {
        ResultModel resultModel = new ResultModel();

        if (StringUtils.isBlank(model.getText())) {
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_EMPTY_TEXT);
            return resultModel;
        }

        Document doc = FileInitHelper.initWord(resultModel, docFile, "findAndHighlight.doc");
        if (doc == null) {
            return resultModel;
        }
        try {
            wordDemoService.findAndHighlight(resultModel, doc, model.getText());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        } finally {
            doc.close();
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/headerAndFooter")
    @ResponseBody
    public ResultModel headerAndFooter(@RequestPart(value = "docFile", required = false) MultipartFile docFile,
                                       @RequestPart(value = "headerImg", required = false) MultipartFile headerImg,
                                       @RequestPart(value = "footerImg", required = false) MultipartFile footerImg,
                                       @RequestPart("jsonModel") WordHeaderAndFooterRequestModel model) {
        ResultModel resultModel = new ResultModel();

        initHeaderAndFooterModel(model);
        Document doc = FileInitHelper.initWord(resultModel, docFile, "headerAndFooter.doc");
        if (doc == null) {
            return resultModel;
        }
        byte[] headerImgData = FileInitHelper.initImage(resultModel, headerImg, "Header.png");
        if (headerImgData == null) {
            return resultModel;
        }
        byte[] footerImgData = FileInitHelper.initImage(resultModel, footerImg, "Footer.png");
        if (footerImgData == null) {
            return resultModel;
        }

        try {
            wordDemoService.addHeaderAndFooter(resultModel, doc, headerImgData, footerImgData, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        } finally {
            doc.close();
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/addTable")
    @ResponseBody
    public ResultModel addTable(@RequestPart("jsonModel") WordTableRequestModel model) {
        ResultModel resultModel = new ResultModel();
        initTableModel(model);
        try {
            wordDemoService.addTable(resultModel, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/mailMerge")
    @ResponseBody
    public ResultModel mailMerge(@RequestPart("jsonModel") WordMailMergeRequestModel model) {
        ResultModel resultModel = new ResultModel();

        try {
            wordDemoService.mailMerge(resultModel, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/waterMark")
    @ResponseBody
    public ResultModel addWaterMark(@RequestPart(value = "docFile", required = false) MultipartFile docFile,
                                    @RequestPart(value = "waterMarkImg", required = false) MultipartFile waterMarkImg,
                                    @RequestPart("jsonModel") WordAddWaterMarkRequestModel model) {
        ResultModel resultModel = new ResultModel();

        initWaterMarkModel(model);
        Document doc = FileInitHelper.initWord(resultModel, docFile, "waterMark.doc");
        if (doc == null) {
            return resultModel;
        }
        byte[] imgData = FileInitHelper.initImage(resultModel, waterMarkImg, "ImageWatermark.png");
        if (imgData == null) {
            return resultModel;
        }
        try {
            wordDemoService.addWaterMark(resultModel, doc, imgData, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        } finally {
            doc.close();
        }
        return resultModel;
    }


    @RequestMapping(value = "/mailMergeTempDownload")
    public void mailMergeTempDownload() throws IOException {
        ServletRequestAttributes requestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = requestAttributes.getResponse();
        CommonHelper.downloadFile(response, Static.WORD_TEMPLATES_FILE_PATH, "Template.doc", false);
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getConvertTypeOptions")
    public ResultModel getConvertTypeOptions() {
        ResultModel resultModel = new ResultModel();
        wordDemoService.getConvertTypeOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getPageSizeOptions")
    public ResultModel getPageSizeOptions() {
        ResultModel resultModel = new ResultModel();
        wordDemoService.getPageSizeOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getTextAlignmentOptions")
    public ResultModel getTextAlignmentOptions() {
        ResultModel resultModel = new ResultModel();
        wordDemoService.getTextAlignmentOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getTableData")
    public ResultModel getTableData() {
        ResultModel resultModel = new ResultModel();
        wordDemoService.getTableData(resultModel);
        return resultModel;
    }

    /**
     * *****************************************************************************************
     */

    private void initWaterMarkModel(WordAddWaterMarkRequestModel model) {
        if (model.getTextSize() == 0) {
            model.setTextSize(80);
        }
        if (StringUtils.isBlank(model.getTextFont())) {
            model.setTextFont("Arial");
        }
        if (StringUtils.isBlank(model.getWaterMarkText())) {
            model.setWaterMarkText("E-iceblue");
        }
        //如果颜色值为空 或者解析颜色错误 那么默认字体颜色为黑色
        if (StringUtils.isBlank(model.getTextColor()) || CommonHelper.String2Color(model.getTextColor()) == null) {
            model.setTextColor("#000000");
        }
        if (StringUtils.isBlank(model.getWaterMarkType())) {
            model.setWaterMarkType("Text");
        }
    }

    /**
     * 检查并初始化所有颜色值
     *
     * @param model
     */
    private void initTableModel(WordTableRequestModel model) {
        if (StringUtils.isBlank(model.getHeaderBackColor()) || CommonHelper.String2Color(model.getHeaderBackColor()) == null) {
            model.setHeaderBackColor("#4bacc6");
        }
        if (StringUtils.isBlank(model.getRowBackColor()) || CommonHelper.String2Color(model.getRowBackColor()) == null) {
            model.setRowBackColor("#d2eaf1");
        }
        if (StringUtils.isBlank(model.getRowBackColor()) || CommonHelper.String2Color(model.getRowBackColor()) == null) {
            model.setRowBackColor("#d2eaf1");
        }
        //默认不使用交替色
        if (StringUtils.isBlank(model.getAlternationBackColor()) || CommonHelper.String2Color(model.getAlternationBackColor()) == null) {
            model.setAlternationBackColor("");
        }
        //边框颜色默认黑
        if (StringUtils.isBlank(model.getBorderColor()) || CommonHelper.String2Color(model.getBorderColor()) == null) {
            model.setBorderColor("#000000");
        }
    }

    /**
     * 检查并初始化所有model参数。如果参数为空则设置默认值
     *
     * @param model
     */
    private void initHeaderAndFooterModel(WordHeaderAndFooterRequestModel model) {
        if (StringUtils.isBlank(model.getPageSize())) {
            model.setPageSize("A4");
        }
        if (StringUtils.isBlank(model.getFooterFont())) {
            model.setFooterFont("Arial");
        }
        if (StringUtils.isBlank(model.getHeaderFont())) {
            model.setHeaderFont("Arial");
        }
        if (StringUtils.isBlank(model.getFooterText())) {
            model.setFooterText("");
        }
        if (StringUtils.isBlank(model.getHeaderText())) {
            model.setHeaderText("");
        }
        //如果颜色值为空 或者解析颜色错误 那么默认字体颜色为黑色
        if (StringUtils.isBlank(model.getHeaderTextColor()) || CommonHelper.String2Color(model.getHeaderTextColor()) == null) {
            model.setHeaderTextColor("#000000");
        }
        if (StringUtils.isBlank(model.getFooterTextColor()) || CommonHelper.String2Color(model.getFooterTextColor()) == null) {
            model.setFooterTextColor("#000000");
        }
        //对齐方式 默认靠左
        if (StringUtils.isBlank(model.getHeaderTextAlignment())) {
            model.setHeaderTextAlignment("Left");
        }
        if (StringUtils.isBlank(model.getFooterTextAlignment())) {
            model.setHeaderTextAlignment("Left");
        }
        if (model.getHeaderFontSize() == 0) {
            model.setHeaderFontSize(14);
        }
        if (model.getFooterFontSize() == 0) {
            model.setFooterFontSize(14);
        }
    }


}
