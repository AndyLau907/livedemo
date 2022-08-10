package com.iceblue.livedemo.controller;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.excel.*;
import com.iceblue.livedemo.service.ExcelDemoService;
import com.iceblue.livedemo.utils.FileInitHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.xls.Workbook;
import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RequestMapping(value = "/excel")
@RestController
public class ExcelController {
    @Autowired
    ExcelDemoService excelDemoService;

    @RequestMapping(method = RequestMethod.POST, value = "/conversion")
    @ResponseBody
    public ResultModel conversion(@RequestPart(value = "excelFile", required = false) MultipartFile excelFile, @RequestPart("jsonModel") ExcelConversionModel model) {
        ResultModel resultModel = new ResultModel();

        if (StringUtils.isBlank(model.getConvertTo())) {
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_SELECT_CONVERSION_TYPE);
            return resultModel;
        }
        Workbook workbook = FileInitHelper.initExcel(resultModel, excelFile, "conversion.xlsx");
        if (workbook == null) {
            return resultModel;
        }
        try {
            excelDemoService.excelConversion(resultModel, workbook, model.getConvertTo());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/addChart")
    @ResponseBody
    public ResultModel addChart(@RequestPart("jsonModel") ExcelChartModel model) {
        ResultModel resultModel = new ResultModel();
        initChartModel(model);
        try {
            excelDemoService.excelChart(resultModel, model.getChartName(), model.getExcelVersion());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/createPivotTable")
    @ResponseBody
    public ResultModel createPivotTable(@RequestPart("jsonModel") ExcelPivotTableModel model) {
        ResultModel resultModel = new ResultModel();
        initPivotTableModel(model);
        try {
            excelDemoService.createPivotTable(resultModel, model.getExcelVersion(), model.getAverageName(), model.getSumName(), model.getMaxName(), model.getMinName());
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/calculateFormulas")
    @ResponseBody
    public ResultModel calculateFormulas(@RequestPart("jsonModel") ExcelCalculateFormulasModel model) {
        ResultModel resultModel = new ResultModel();
        initFormulasModel(model);
        try {
            excelDemoService.calculateFormulas(resultModel, model);
        } catch (Exception e) {
            e.printStackTrace();
            resultModel.setValid(false);
            resultModel.setMessage(Static.MESSAGE_FAILED_PROCESS);
            return resultModel;
        }
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.POST, value = "/useWorkbookDesigner")
    @ResponseBody
    public ResultModel useWorkbookDesigner(@RequestPart("jsonModel") ExcelWorkbookDesignerModel model) {
        ResultModel resultModel = new ResultModel();
        if (StringUtils.isBlank(model.getExcelVersion())) {
            model.setExcelVersion("Excel 2010 Workbook");
        }
        try {
            excelDemoService.useWorkbookDesigner(resultModel, model.getExcelVersion());
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
        excelDemoService.getConvertOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getExcelChartOptions")
    public ResultModel getExcelChartOptions() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.getExcelChartOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getExcelVersionOptions")
    public ResultModel getExcelVersionOptions() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.getExcelVersionOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getPivotTableOptions")
    public ResultModel getPivotTableOptions() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.getPivotTableOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getExcelMathematicOptions")
    public ResultModel getExcelMathematicOptions() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.getExcelMathematicOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getExcelLogicOptions")
    public ResultModel getExcelLogicOptions() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.getExcelLogicOptions(resultModel);
        return resultModel;
    }

    /*@RequestMapping(method = RequestMethod.GET, value = "/getExcelLogicBoolOptions")
    public ResultModel getExcelLogicBoolOptions() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.getExcelLogicBoolOptions(resultModel);
        return resultModel;
    }*/

    @RequestMapping(method = RequestMethod.GET, value = "/getExcelSimpleOptions")
    public ResultModel getExcelSimpleOptions() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.getExcelSimpleOptions(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/showChartData")
    public ResultModel showChartData() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.showChartData(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/showPivotTableData")
    public ResultModel showPivotTableData() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.showPivotTableData(resultModel);
        return resultModel;
    }

    @RequestMapping(method = RequestMethod.GET, value = "/showWorkbookDesignerData")
    public ResultModel showWorkbookDesignerData() {
        ResultModel resultModel = new ResultModel();
        excelDemoService.showWorkbookDesignerData(resultModel);
        return resultModel;
    }

    /*******************************************************************************************/
    public void initFormulasModel(ExcelCalculateFormulasModel model) {
        if (StringUtils.isBlank(model.getMathFunction())) {
            model.setMathFunction("ABS");
        }
        if (StringUtils.isBlank(model.getLogicFunction())) {
            model.setLogicFunction("AND");
        }
        if (StringUtils.isBlank(model.getSimpleExpression())) {
            model.setMathFunction("SUM");
        }
        if (StringUtils.isBlank(model.getMidText())) {
            model.setMathFunction("Hello world!");
        }
        if (StringUtils.isBlank(model.getExcelVersion())) {
            model.setExcelVersion("Excel 2010 Workbook");
        }
    }

    public void initChartModel(ExcelChartModel model) {
        if (StringUtils.isBlank(model.getExcelVersion())) {
            model.setExcelVersion("Excel 2010 Workbook");
        }
        if (StringUtils.isBlank(model.getChartName())) {
            model.setChartName("ColumnClustered");
        }
    }

    public void initPivotTableModel(ExcelPivotTableModel model) {
        if (StringUtils.isBlank(model.getAverageName())) {
            model.setAverageName("Sales");
        }
        if (StringUtils.isBlank(model.getSumName())) {
            model.setSumName("Sales");
        }
        if (StringUtils.isBlank(model.getMaxName())) {
            model.setMaxName("Sales");
        }
        if (StringUtils.isBlank(model.getMinName())) {
            model.setMinName("Sales");
        }
        if (StringUtils.isBlank(model.getExcelVersion())) {
            model.setExcelVersion("Excel 2010 Workbook");
        }
    }
}
