package com.iceblue.livedemo.service;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.TypeOptionsModel;
import com.iceblue.livedemo.model.excel.ExcelCalculateFormulasModel;
import com.iceblue.livedemo.model.excel.ExcelChartDataModel;
import com.iceblue.livedemo.model.excel.ExcelPivotTableDataModel;
import com.iceblue.livedemo.model.excel.ExcelWorkbookDesignerDataModel;
import com.iceblue.livedemo.utils.CommonHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.data.table.DataTable;
import com.spire.xls.*;
import com.spire.xls.core.IPivotField;
import com.spire.xls.core.IXLSRange;
import com.spire.xls.core.spreadsheet.HTMLOptions;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @author Leiq
 * @program: LiveDemo
 * @description:
 * @date 2021-10-15 17:38:58
 */

@Service
public class ExcelDemoService {

    /**
     * excel的转换
     */
    public void excelConversion(ResultModel resultModel, Workbook workbook, String convertTo) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID();
        switch (convertTo) {
            case "PDF":
                outputFile += ".pdf";
                workbook.getConverterSetting().setSheetFitToPage(true);
                workbook.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PDF);
                break;
            case "IMAGE":
                BufferedImage[] images=workbook.saveAsImage(300,300);
                outputFile=CommonHelper.packageImages(images, "png");
                break;
            case "HTML":
                outputFile = packageHtml(workbook);
                break;
            case "TIFF":
                outputFile += ".tiff";
                workbook.saveToTiff(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "XPS":
                outputFile += ".xps";
                workbook.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.XPS);
                break;
        }
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    /**
     * Excel 添加图表chart
     */
    public void excelChart(ResultModel resultModel, String chartName, String excelVersion) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + getExcelType(excelVersion);
        Workbook workbook = new Workbook();
        workbook.loadFromFile(Static.Excel_TEMPLATES_FILE_PATH + "Chart.xlsx");
        Worksheet sheet = workbook.getWorksheets().get(0);
        Chart chart = sheet.getCharts().add();

        chart.setChartType(getChartType(chartName));
        chart.setDataRange(sheet.getCellRange("A1:C7"));
        chart.setSeriesDataFromRange(false);

        //set position of chart
        chart.setLeftColumn(4);
        chart.setTopRow(2);
        chart.setRightColumn(12);
        chart.setBottomRow(22);

        chart.setChartTitle("Sales market by country");
        chart.getChartTitleArea().isBold(true);
        chart.getChartTitleArea().setSize(12);

        chart.getPrimarySerieAxis().setTitle("Country");
        chart.getPrimarySerieAxis().getFont().isBold(true);
        chart.getPrimarySerieAxis().getTitleArea().isBold(true);

        chart.getPrimarySerieAxis().setTitle("Sales(in Dollars)");
        chart.getPrimarySerieAxis().hasMajorGridLines(false);
        chart.getPrimarySerieAxis().getTitleArea().setTextRotationAngle(90);
        //chart.getPrimarySerieAxis().setMinValue(1000);
        chart.getPrimarySerieAxis().getTitleArea().isBold(true);

        chart.getPlotArea().getFill().setFillType(ShapeFillType.SolidColor);
        chart.getPlotArea().getFill().setForeKnownColor(ExcelColors.White);

        for (int i = 0; i < chart.getSeries().getCount(); i++) {
            chart.getSeries().get(i).getFormat().getOptions().isVaryColor(true);
            chart.getSeries().get(i).getDataPoints().getDefaultDataPoint().getDataLabels().hasValue(true);
        }
        chart.getLegend().setPosition(LegendPositionType.Top);

        CellStyle oddStyle = workbook.getStyles().addStyle("oddStyle");
        oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
        oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
        oddStyle.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
        oddStyle.setKnownColor(ExcelColors.LightGreen1);

        CellStyle evenStyle = workbook.getStyles().addStyle("evenStyle");
        evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
        evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
        evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
        evenStyle.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
        evenStyle.setKnownColor(ExcelColors.LightTurquoise);

        for (int i = 0; i < sheet.getAllocatedRange().getRows().length; i++) {
            CellRange[] ranges = sheet.getAllocatedRange().getRows();
            if (ranges[i].getRow() != 0) {
                if (ranges[i].getRow() % 2 == 0) {
                    ranges[i].setCellStyleName(evenStyle.getName());
                } else {
                    ranges[i].setCellStyleName(oddStyle.getName());
                }
            }
        }

        //Sets header style
        CellStyle styleHeader = workbook.getStyles().addStyle("headerStyle");
        styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);
        styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin);
        styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin);
        styleHeader.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin);
        styleHeader.setVerticalAlignment(VerticalAlignType.Center);
        styleHeader.setKnownColor(ExcelColors.Green);
        styleHeader.getFont().setKnownColor(ExcelColors.White);
        styleHeader.getFont().isBold(true);
        styleHeader.setHorizontalAlignment(HorizontalAlignType.Center);

        for (int i = 0; i < sheet.getRows()[0].getCount(); i++) {
            CellRange range = sheet.getRows()[0];
            range.setCellStyleName(styleHeader.getName());
        }
        sheet.getColumns()[sheet.getAllocatedRange().getLastColumn() - 1].getStyle().setNumberFormat("\"$\"#,##0");
        sheet.getColumns()[sheet.getAllocatedRange().getLastColumn() - 2].getStyle().setNumberFormat("\"$\"#,##0");
        sheet.getRows()[0].getStyle().setNumberFormat("General");
        sheet.getAllocatedRange().autoFitColumns();
        sheet.getAllocatedRange().autoFitRows();
        sheet.getRows()[0].setRowHeight(20);

        workbook.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, getFileFormat(excelVersion));
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    /**
     * excel创建Pivot Table
     */
    public void createPivotTable(ResultModel resultModel, String excelVersion, String averageName, String sumName, String maxName, String minName) throws IOException {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + getExcelType(excelVersion);
        Workbook workbook = new Workbook();
        String filePath = Static.Excel_TEMPLATES_FILE_PATH + "PartSalesInfo.xlsx";
        workbook.loadFromFile(filePath, ExcelVersion.Version2007);
        Worksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.setName("Data Source");
        Worksheet worksheet1 = workbook.createEmptySheet();
        worksheet1.setName("Pivot Table");
        CellRange dataRange = worksheet.getRange().get("A1:G17");
        PivotCache cache = workbook.getPivotCaches().add(dataRange);
        PivotTable pt = worksheet1.getPivotTables().add("Pivot Table", worksheet.getRange().get("A1"), cache);
        IPivotField r1 = pt.getPivotFields().get("Vendor No");

        r1.setAxis(AxisTypes.Row);
        pt.getOptions().setRowHeaderCaption("Vendor No");
        IPivotField r2 = pt.getPivotFields().get("Name");
        r2.setAxis(AxisTypes.Row);
        pt.getDataFields().add(pt.getPivotFields().get(averageName), "Average of", SubtotalTypes.Average);
        pt.getDataFields().add(pt.getPivotFields().get(sumName), "SUM of", SubtotalTypes.Sum);
        pt.getDataFields().add(pt.getPivotFields().get(maxName), "Max of", SubtotalTypes.Max);
        pt.getDataFields().add(pt.getPivotFields().get(minName), "Min of", SubtotalTypes.Min);
        pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleMedium12);
        pt.calculateData();
        worksheet1.getAllocatedRange().autoFitColumns();
        workbook.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, getFileFormat(excelVersion));
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    /**
     * 计算公式
     */
    public void calculateFormulas(ResultModel resultModel, ExcelCalculateFormulasModel model) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + getExcelType(model.getExcelVersion());
        Workbook workbook = new Workbook();
        workbook.createEmptySheets(1);
        Worksheet worksheet = workbook.getWorksheets().get(0);
        CalculateFormuLas(workbook, model.getMathFunction(), Float.toString(model.getMathData()), model.getLogicFunction(), model.isLogicValueOne(),
                model.isLogicValueTwo(), model.getSimpleExpression(), Integer.toString(model.getsData()), Integer.toString(model.getsData()), model.getMidText(),
                Integer.toString(model.getStartNumber()), Integer.toString(model.getNumberChart()));
        workbook.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, getFileFormat(model.getExcelVersion()));
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);

    }

    /**
     * WorkbookDesigner的使用
     */
    public void useWorkbookDesigner(ResultModel resultModel, String excelVersion) throws IOException {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + getExcelType(excelVersion);
        Workbook workbook = new Workbook();
        String filePath;
        if (excelVersion.contains("Excel 2007 Workbook") || excelVersion.contains("Excel 2010 Workbook")) {
            filePath = Static.Excel_TEMPLATES_FILE_PATH + "MarkerDesignerSample2007.xlsx";
        } else {
            filePath = Static.Excel_TEMPLATES_FILE_PATH + "MarkerDesignerSample.xls";
        }
        workbook.loadFromFile(filePath);
        Worksheet sheet = workbook.getWorksheets().get(0);
        Worksheet sheet2 = workbook.getWorksheets().get(1);
        sheet.setName("Result");
        sheet2.setName("DataSource");
        sheet2.insertDataTable(getData(), true, 1, 1);
        workbook.getMarkerDesigner().addParameter("Variable1", 1234.5678);
        workbook.getMarkerDesigner().addDataTable("Country", getData());
        workbook.getMarkerDesigner().apply();
        sheet.getAllocatedRange().autoFitRows();
        sheet.getAllocatedRange().autoFitColumns();

        workbook.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, getFileFormat(excelVersion));
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void getConvertOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "excelConvertOptions.xml", TypeOptionsModel.class));

    }

    public void getExcelChartOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "excelChartOptions.xml", TypeOptionsModel.class));
    }

    public void getExcelVersionOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "excelVersionOptions.xml", TypeOptionsModel.class));
    }

    public void getPivotTableOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pivotTableOptions.xml", TypeOptionsModel.class));
    }

    public void getExcelMathematicOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "excelMathematicOptions.xml", TypeOptionsModel.class));
    }

    public void getExcelLogicOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "excelLogicOptions.xml", TypeOptionsModel.class));
    }

    public void getExcelLogicBoolOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "excelLoficBoolOptions.xml", TypeOptionsModel.class));
    }

    public void getExcelSimpleOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "excelSimpleOptions.xml", TypeOptionsModel.class));
    }

    /**
     * excel chart 原始数据
     */
    public void showChartData(ResultModel resultModel) {
        List<ExcelChartDataModel> chartDataModelList = new ArrayList<>();
        Workbook workbook = new Workbook();
        workbook.loadFromFile(Static.Excel_TEMPLATES_FILE_PATH + "Chart.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);
        for (int i = 1; i < worksheet.getRows().length; i++) {
            String country = worksheet.getRows()[i].getCellList().get(0).getValue();
            double jun = worksheet.getRows()[i].getCellList().get(1).getNumberValue();
            double aug = worksheet.getRows()[i].getCellList().get(2).getNumberValue();
            ExcelChartDataModel chartDataModel = new ExcelChartDataModel(country, jun, aug);
            chartDataModelList.add(chartDataModel);
        }
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(chartDataModelList);
    }

    /**
     * excel 透明表的展示数据
     */
    public void showPivotTableData(ResultModel resultModel) {
        List<ExcelPivotTableDataModel> pivotTableDataModelList = new ArrayList<>();
        Workbook workbook = new Workbook();
        String filePath = Static.Excel_TEMPLATES_FILE_PATH + "PartSalesInfo.xlsx";
        workbook.loadFromFile(filePath, ExcelVersion.Version2007);
        Worksheet worksheet = workbook.getWorksheets().get(0);
        for (int i = 1; i < worksheet.getRows().length; i++) {
            String name = worksheet.getRows()[i].getCellList().get(0).getValue();
            int vendorNo = (int) worksheet.getRows()[i].getCellList().get(1).getNumberValue();
            double sales = worksheet.getRows()[i].getCellList().get(2).getNumberValue();
            int onHand = (int) worksheet.getRows()[i].getCellList().get(3).getNumberValue();
            int onOrder = (int) worksheet.getRows()[i].getCellList().get(4).getNumberValue();
            int area = (int) worksheet.getRows()[i].getCellList().get(5).getNumberValue();
            long population = (long) worksheet.getRows()[i].getCellList().get(6).getNumberValue();
            ExcelPivotTableDataModel excelPivotTableDataModel = new ExcelPivotTableDataModel(name, vendorNo, sales, onHand, onOrder, area, population);
            pivotTableDataModelList.add(excelPivotTableDataModel);
        }
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(pivotTableDataModelList);
    }


    /**
     * WorkbookDesigner展示数据
     */
    public void showWorkbookDesignerData(ResultModel resultModel) {
        List<ExcelWorkbookDesignerDataModel> workbokDesignerDatamodelList = new ArrayList<>();
        Workbook workbook = new Workbook();
        String filePath = Static.Excel_TEMPLATES_FILE_PATH + "DatatableSample.xls";
        workbook.loadFromFile(filePath, ExcelVersion.Version97to2003);
        Worksheet worksheet = workbook.getWorksheets().get(0);
        for (int i = 1; i < worksheet.getRows().length; i++) {
            String name = worksheet.getRows()[i].getCellList().get(0).getValue();
            String capital = worksheet.getRows()[i].getCellList().get(1).getValue();
            String continent = worksheet.getRows()[i].getCellList().get(2).getValue();
            int area = (int) worksheet.getRows()[i].getCellList().get(3).getNumberValue();
            long population = (long) worksheet.getRows()[i].getCellList().get(4).getNumberValue();
            ExcelWorkbookDesignerDataModel workbokDesignerDatamodel = new ExcelWorkbookDesignerDataModel(name, capital, continent, area, population);
            workbokDesignerDatamodelList.add(workbokDesignerDatamodel);
        }
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);

        resultModel.setData(workbokDesignerDatamodelList);
    }


    /**
     * 打包HTML文件
     */
    private String packageHtml(Workbook workbook) throws Exception {
        String ID = UUID.randomUUID().toString();
        String outputFile = Static.OUTPUT_FILE_START + ID + ".zip";
        String diretoryPath = Static.OUTPUT_FILE_PATH + ID;
        File htmlFolder = new File(diretoryPath);
        if (!htmlFolder.exists()) {
            htmlFolder.mkdir();
        }
        for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
            HTMLOptions options = new HTMLOptions();
            options.setImageEmbedded(true);
            String htmlPath = htmlFolder + "/" + "output"+(i+1)+".html";
            Worksheet sheet=workbook.getWorksheets().get(i);
            sheet.saveToHtml(htmlPath, options);
        }
        FileOutputStream fos = new FileOutputStream(Static.OUTPUT_FILE_PATH+outputFile);
        CommonHelper.toZip(diretoryPath, fos, true);
        CommonHelper.deleteDirectory(diretoryPath);
        return outputFile;
    }

    /**
     * 获取excel文件的后缀名
     */
    private String getExcelType(String excelVersion) {
        String type = ".xls";
        switch (excelVersion) {
            case "Excel 97to2003 Workbook":
                type = ".xls";
                break;
            case "Excel 2007 Workbook":
            case "Excel 2010 Workbook":
            case "Excel 2013 Workbook":
            case "Excel 2016 Workbook":
                type = ".xlsx";
                break;
            case "Excel 2007 Binary Workbook":
            case "Excel 2010 Binary Workbook":
                type = ".xlsb";
                break;
        }
        return type;
    }

    /**
     * 判断excel文件版本信息
     */
    private FileFormat getFileFormat(String excelVersion) {
        FileFormat fileFormat = FileFormat.Version97to2003;
        switch (excelVersion) {
            case "Excel 97to2003 Workbook":
                fileFormat = FileFormat.Version97to2003;
                break;
            case "Excel 2007 Workbook":
                fileFormat = FileFormat.Version2007;
                break;
            case "Excel 2010 Workbook":
                fileFormat = FileFormat.Version2010;
                break;
            case "Excel 2007 Binary Workbook":
                fileFormat = FileFormat.Xlsb2007;
                break;
            case "Excel 2010 Binary Workbook":
                fileFormat = FileFormat.Xlsb2010;
                break;
            case "Excel 2013 Workbook":
                fileFormat=FileFormat.Version2013;
                break;
            case "Excel 2016 Workbook":
                fileFormat=FileFormat.Version2016;
                break;
        }
        return fileFormat;
    }

    /**
     * 获取chart类型
     */

    private ExcelChartType getChartType(String chartName) {
        ExcelChartType chartType = ExcelChartType.ColumnClustered;
        switch (chartName) {
            case "ColumnClustered":
                chartType = ExcelChartType.ColumnClustered;
                break;
            case "ColumnStacked":
                chartType = ExcelChartType.ColumnStacked;
                break;
            case "Column3DClustered":
                chartType = ExcelChartType.Column3DClustered;
                break;
            case "Column3DStacked":
                chartType = ExcelChartType.Column3DStacked;
                break;
            case "Column3D":
                chartType = ExcelChartType.Column3D;
                break;
            case "BarClustered":
                chartType = ExcelChartType.BarClustered;
                break;
            case "BarStacked":
                chartType = ExcelChartType.BarStacked;
                break;
            case "Bar3DClustered":
                chartType = ExcelChartType.Bar3DClustered;
                break;
            case "Bar3DStacked":
                chartType = ExcelChartType.Bar3DStacked;
                break;
            case "Line":
                chartType = ExcelChartType.Line;
                break;
            case "LineStacked":
                chartType = ExcelChartType.LineStacked;
                break;
            case "LineMarkers":
                chartType = ExcelChartType.LineMarkers;
                break;
            case "LineMarkersStacked":
                chartType = ExcelChartType.LineMarkersStacked;
                break;
            case "Line3D":
                chartType = ExcelChartType.Line3D;
                break;
            case "Pie":
                chartType = ExcelChartType.Pie;
                break;
            case "Pie3D":
                chartType = ExcelChartType.Pie3D;
                break;
            case "PieOfPie":
                chartType = ExcelChartType.PieOfPie;
                break;
            case "PieExploded":
                chartType = ExcelChartType.PieExploded;
                break;
            case "Pie3DExploded":
                chartType = ExcelChartType.Pie3DExploded;
                break;
            case "PieBar":
                chartType = ExcelChartType.PieBar;
                break;
            case "ScatterMarkers":
                chartType = ExcelChartType.ScatterMarkers;
                break;
            case "ScatterSmoothedLineMarkers":
                chartType = ExcelChartType.ScatterSmoothedLineMarkers;
                break;
            case "ScatterSmoothedLine":
                chartType = ExcelChartType.ScatterSmoothedLine;
                break;
            case "ScatterLineMarkers":
                chartType = ExcelChartType.ScatterLineMarkers;
                break;
            case "ScatterLine":
                chartType = ExcelChartType.ScatterLine;
                break;
            case "Area":
                chartType = ExcelChartType.Area;
                break;
            case "AreaStacked":
                chartType = ExcelChartType.AreaStacked;
                break;
            case "Area3D":
                chartType = ExcelChartType.Area3D;
                break;
            case "Area3DStacked":
                chartType = ExcelChartType.Area3DStacked;
                break;
            case "Doughnut":
                chartType = ExcelChartType.Doughnut;
                break;
            case "DoughnutExploded":
                chartType = ExcelChartType.DoughnutExploded;
                break;
            case "Radar":
                chartType = ExcelChartType.Radar;
                break;
            case "RadarMarkers":
                chartType = ExcelChartType.RadarMarkers;
                break;
            case "RadarFilled":
                chartType = ExcelChartType.RadarFilled;
                break;
            case "Surface3D":
                chartType = ExcelChartType.Surface3D;
                break;
            case "Surface3DNoColor":
                chartType = ExcelChartType.Surface3DNoColor;
                break;
            case "SurfaceContour":
                chartType = ExcelChartType.SurfaceContour;
                break;
            case "SurfaceContourNoColor":
                chartType = ExcelChartType.SurfaceContourNoColor;
                break;
            case "Bubble":
                chartType = ExcelChartType.Bubble;
                break;
            case "Bubble3D":
                chartType = ExcelChartType.Bubble3D;
                break;
            case "CylinderClustered":
                chartType = ExcelChartType.CylinderClustered;
                break;
            case "CylinderStacked":
                chartType = ExcelChartType.CylinderStacked;
                break;
            case "CylinderBarClustered":
                chartType = ExcelChartType.CylinderBarClustered;
                break;
            case "CylinderBarStacked":
                chartType = ExcelChartType.CylinderBarStacked;
                break;
            case "Cylinder3DClustered":
                chartType = ExcelChartType.Cylinder3DClustered;
                break;
            case "ConeClustered":
                chartType = ExcelChartType.ConeClustered;
                break;
            case "ConeStacked":
                chartType = ExcelChartType.ConeStacked;
                break;
            case "ConeBarClustered":
                chartType = ExcelChartType.ConeBarClustered;
                break;
            case "ConeBarStacked":
                chartType = ExcelChartType.ConeBarStacked;
                break;
            case "Cone3DClustered":
                chartType = ExcelChartType.Cone3DClustered;
                break;
            case "PyramidClustered":
                chartType = ExcelChartType.PyramidClustered;
                break;
            case "PyramidStacked":
                chartType = ExcelChartType.PyramidStacked;
                break;
            case "PyramidBarClustered":
                chartType = ExcelChartType.PyramidBarClustered;
                break;
            case "PyramidBarStacked":
                chartType = ExcelChartType.PyramidBarStacked;
                break;
            case "Pyramid3DClustered":
                chartType = ExcelChartType.Pyramid3DClustered;
                break;
        }

        return chartType;
    }

    /**
     * 计算公式
     */
    private void CalculateFormuLas(Workbook workbook, String mathFunction, String mathData, String logicFunction,
                                   boolean logicValueOne, boolean logicValueTwo, String simpleExpression, String fData,
                                   String sData, String midText, String startNumber, String numberChart) {
        Worksheet worksheet = workbook.getWorksheets().get(0);
        int currentRow = 1;
        String currentFormula = null;
        Object formulaResult = null;
        String value = null;
        worksheet.setColumnWidth(1, 32);
        worksheet.setColumnWidth(2, 16);
        worksheet.setColumnWidth(3, 16);
        CellRange range = worksheet.getRange().get("A1");
        worksheet.getRange().get(currentRow, 1).setValue("Formulas");
        worksheet.getRange().get(currentRow, 2).setValue("Results");
        worksheet.getRange().get(currentRow, 2).setHorizontalAlignment(HorizontalAlignType.Right);
        range = worksheet.getRange().get(currentRow, 1, currentRow, 2);
        range.getStyle().getFont().isBold(true);
        range.getStyle().setKnownColor(ExcelColors.LightGreen1);
        range.getStyle().setFillPattern(ExcelPatternType.Solid);
        range.getStyle().getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Medium);
        currentFormula = "=" + mathFunction + "(" + mathData + ")";
        worksheet.getRange().get(++currentRow, 1).setText(currentFormula);
        formulaResult = workbook.calculateFormulaValue(currentFormula);
        value = formulaResult.toString();
        worksheet.getRange().get(currentRow, 2).setValue(value);
        if (mathFunction == "NOT") {
            currentFormula = "=" + logicFunction + "(" + logicValueOne + ")";
            logicValueTwo = false;
        } else {
            currentFormula = "=" + logicFunction + "(" + logicValueOne + "," + logicValueTwo + ")";
        }
        worksheet.getRange().get(++currentRow, 1).setText(currentFormula);
        formulaResult = workbook.calculateFormulaValue(currentFormula);
        value = formulaResult.toString();
        worksheet.getRange().get(currentRow, 2).setValue(value);
        worksheet.getRange().get(currentRow, 2).setHorizontalAlignment(HorizontalAlignType.Right);
        if (simpleExpression == "SUBTRACT") {
            currentFormula = "=SUM(" + fData + "," + "-" + sData + ")";
        } else {
            currentFormula = "=" + simpleExpression + "(" + fData + "," + sData + ")";
        }
        worksheet.getRange().get(++currentRow, 1).setText(currentFormula);
        formulaResult = workbook.calculateFormulaValue(currentFormula);
        value = formulaResult.toString();
        worksheet.getRange().get(currentRow, 2).setValue(value);
        currentFormula = "MID(\"" + midText + "\"," + startNumber + "," + numberChart + ")";
        worksheet.getRange().get(++currentRow, 1).setText(currentFormula);
        formulaResult = workbook.calculateFormulaValue(currentFormula);
        value = formulaResult.toString();
        worksheet.getRange().get(currentRow, 2).setValue(value);
        worksheet.getRange().get(currentRow, 2).setHorizontalAlignment(HorizontalAlignType.Right);
        currentFormula = "=RAND()";
        worksheet.getRange().get(++currentRow, 1).setText(currentFormula);
        formulaResult = workbook.calculateFormulaValue(currentFormula);
        value = formulaResult.toString();
        worksheet.getRange().get(currentRow, 2).setValue(value);
    }

    /**
     * 将excel数据转到dataTable，并返回DataTable
     */
    private DataTable getData() throws IOException {
        Workbook workbook = new Workbook();
        String filePath = Static.Excel_TEMPLATES_FILE_PATH + "DatatableSample.xls";
        workbook.loadFromFile(filePath);
        Worksheet sheet = workbook.getWorksheets().get(0);
        return sheet.exportDataTable();
    }
}
