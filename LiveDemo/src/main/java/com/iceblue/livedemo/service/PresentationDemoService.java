package com.iceblue.livedemo.service;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.TypeOptionsModel;
import com.iceblue.livedemo.model.powerpoint.PresentationAddWatermarkModel;
import com.iceblue.livedemo.model.powerpoint.ReportModel;
import com.iceblue.livedemo.utils.CommonHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.data.table.DataColumn;
import com.spire.data.table.DataRow;
import com.spire.data.table.DataTable;
import com.spire.pdf.tables.table.DataTypes;
import com.spire.presentation.*;
import com.spire.presentation.charts.ChartType;
import com.spire.presentation.charts.IChart;
import com.spire.presentation.drawing.BackgroundType;
import com.spire.presentation.drawing.FillFormatType;
import com.spire.presentation.drawing.IImageData;
import com.spire.presentation.drawing.PictureFillType;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.*;
import java.util.List;

/**
 * @author Leiq
 * @program: LiveDemo
 * @description:
 * @date 2021-10-15 11:19:59
 */

@Service
public class PresentationDemoService {

    /**
     * powerpoint的转换
     */
    public void preaentionConversion(ResultModel resultModel, Presentation presentation, String convertTo) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID();
        switch (convertTo) {
            case "PDF":
                outputFile += ".pdf";
                presentation.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PDF);
                break;
            case "HTML":
                outputFile = packageHtml(presentation);
                break;
            case "IMAGE":
                BufferedImage[] images = new BufferedImage[presentation.getSlides().getCount()];
                for (int i = 0; i < presentation.getSlides().getCount(); i++) {
                    images[i] = presentation.getSlides().get(i).saveAsImage();
                }
                outputFile=CommonHelper.packageImages(images, "png");
                break;
            case "XPS":
                outputFile += ".xps";
                presentation.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.XPS);
                break;
            case "TIFF":
                outputFile += ".tiff";
                presentation.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.TIFF);
                break;
        }
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    /**
     * 查找高亮显示或替换文本
     */
    public void findOrReplace(ResultModel resultModel, Presentation presentation, String typeChance, String findText, String newText) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pptx";
        switch (typeChance) {
            case "FIND":
                findText(presentation, findText);
                break;
            case "REPLACE":
                replaceText(presentation, findText, newText);
                break;
        }
        presentation.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PPTX_2013);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    /**
     * 添加文本或图片
     */
    public void addTextOrImage(ResultModel resultModel, Presentation presentation, String addType, String text, String fontName, String colorText, float fontSize, byte[] imageData) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pptx";
        switch (addType) {
            case "TEXT":
                addText(presentation, text, fontName, CommonHelper.String2Color(colorText), fontSize);
                break;
            case "IMAGE":
                addImage(presentation, imageData);
                break;
        }
        presentation.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PPTX_2013);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }


    /**
     * 添加水印
     */

    public void addTextOrImageWatermark(ResultModel resultModel, Presentation presentation,PresentationAddWatermarkModel model, byte[] imageData) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pptx";
        switch (model.getWatermarkType().toUpperCase()) {
            case "TEXT":
                addTextWatermark(presentation, model);
                break;
            case "IMAGE":
                addImageWatermark(presentation, imageData);
                break;
        }
        presentation.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PPTX_2013);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    /**
     * 添加图表chart
     */
    public void charts(ResultModel resultModel, String chartName) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pptx";
        Presentation presentation = new Presentation();
        Rectangle2D rect1 = new Rectangle2D.Double(90, 100, 550, 320);
        IChart chart = presentation.getSlides().get(0).getShapes().appendChart(getChartType(chartName), rect1, false);
        //chart title
        chart.getChartTitle().getTextProperties().setText("Chart");
        chart.getChartTitle().getTextProperties().isCentered(true);
        chart.getChartTitle().setHeight(30);
        chart.hasTitle(true);
        DataTable dataTable = getDataTable();
        insertDatatableToChart(chart, dataTable);

        ChartType chartType = getChartType(chartName);
        //set series label
        chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "D1"));

        if (chartType.getName().contains("Scatter") || chartType.getName().contains("Bubble")){
            chart.getSeries().get(0).setXValues(chart.getChartData().get("A2", "A7"));
            chart.getSeries().get(0).setYValues(chart.getChartData().get("B2", "B7"));
            chart.getSeries().get(1).setYValues(chart.getChartData().get("C2", "C7"));
            chart.getSeries().get(2).setYValues(chart.getChartData().get("D2", "D7"));
            if (chartType.getName().contains("Bubble")){
                for (int i = 0; i < 3;i++){
                    chart.getSeries().get(i).getBubbles().add(1);
                    chart.getSeries().get(i).getBubbles().add(4);
                    chart.getSeries().get(i).getBubbles().add(3);
                    chart.getSeries().get(i).getBubbles().add(4);
                    chart.getSeries().get(i).getBubbles().add(2);
                    chart.getSeries().get(i).getBubbles().add(9);
                }
            }
        }else {
            //set category label
            chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "A7"));

            //set values for series
            chart.getSeries().get(0).setValues(chart.getChartData().get("B2", "B7"));
            chart.getSeries().get(1).setValues(chart.getChartData().get("C2", "C7"));
            chart.getSeries().get(2).setValues(chart.getChartData().get("D2", "D7"));
            if (chartType.getName().contains("3D")){
                chart.getRotationThreeD().setXDegree(10);
                chart.getRotationThreeD().setYDegree(10);
            }
        }

        presentation.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PPTX_2013);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }


    public void getConvertTypeOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pptConvertOptions.xml", TypeOptionsModel.class));

    }
    public void getRotateOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pptRotateOptions.xml", TypeOptionsModel.class));
    }

    public void getPptChartOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pptChartOptions.xml", TypeOptionsModel.class));
    }

    public void getChartData(ResultModel resultModel) {
        List<ReportModel> dataList = CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pptChartData.xml", ReportModel.class);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(dataList);
    }


    private String packageHtml(Presentation presentation) throws Exception {
        String ID = UUID.randomUUID().toString();
        String outputFile = Static.OUTPUT_FILE_START + ID + ".zip";
        String htmlFile = "output.html";
        String diretoryPath = Static.OUTPUT_FILE_PATH + ID;
        File htmlFolder = new File(Static.OUTPUT_FILE_PATH + ID);
        if (!htmlFolder.exists()) {
            htmlFolder.mkdir();
        }
        presentation.saveToFile(diretoryPath + "/" + htmlFile, FileFormat.HTML);
        FileOutputStream fos = new FileOutputStream(Static.OUTPUT_FILE_PATH+outputFile);
        CommonHelper.toZip(diretoryPath, fos, true);
        CommonHelper.deleteDirectory(diretoryPath);
        return outputFile;
    }

    private void findText(Presentation presentation, String findText) {
        for (int i = 0; i < presentation.getSlides().getCount(); i++) {
            ISlide slide = presentation.getSlides().get(i);
            for (int j = 0; j < slide.getShapes().getCount(); j++) {
                IAutoShape shape = (IAutoShape) slide.getShapes().get(j);
                TextHighLightingOptions options = new TextHighLightingOptions();
                options.setWholeWordsOnly(true);
                options.setCaseSensitive(true);
                shape.getTextFrame().highLightText(findText, Color.yellow, options);
            }
        }
    }

    private void replaceText(Presentation presentation, String oldText, String newText) {
        for (int i = 0; i < presentation.getSlides().getCount(); i++) {
            ISlide slide = presentation.getSlides().get(i);
            Map<String, String> map = new HashMap<>();
            map.put(oldText, newText);
            replaceTags(slide, map);
        }
    }


    private void addText(Presentation presentation, String text, String fontName, Color color, float fontSize) throws Exception {
        IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(50, 70, 620, 150));
        shape.getFill().setFillType(FillFormatType.NONE);
        shape.getShapeStyle().getLineColor().setColor(Color.white);
        shape.getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.LEFT);
        shape.getTextFrame().getParagraphs().get(0).setIndent(50);
        shape.getTextFrame().getParagraphs().get(0).setLineSpacing(150);
        shape.getTextFrame().setText(text);
        shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setLatinFont(new TextFont(fontName));
        shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setFontHeight(fontSize);
        shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
        shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().getSolidColor().setColor(color);
    }

    private void addImage(Presentation presentation, byte[] imageData) throws Exception {
        Rectangle2D.Double rect1 = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 280, 140, 120, 120);
        ByteArrayInputStream bais = new ByteArrayInputStream(imageData);
        IEmbedImage image = presentation.getSlides().get(0).getShapes().appendEmbedImage(ShapeType.RECTANGLE, bais, rect1);
        image.getLine().setFillType(FillFormatType.NONE);
    }


    private static void replaceTags(ISlide pSlide, Map<String, String> tagValues) {
        for (int i = 0; i < pSlide.getShapes().getCount(); i++) {
            IShape curShape = pSlide.getShapes().get(i);
            if (curShape instanceof IAutoShape) {
                for (int j = 0; j < ((IAutoShape) curShape).getTextFrame().getParagraphs().getCount(); j++) {
                    ParagraphEx tp = ((IAutoShape) curShape).getTextFrame().getParagraphs().get(j);
                    for (Map.Entry<String, String> entry : tagValues.entrySet()) {
                        String mapKey = entry.getKey();
                        String mapValue = entry.getValue();
                        if (tp.getText().contains(mapKey)) {
                            tp.setText(tp.getText().replace(mapKey, mapValue));
                        }
                    }
                }
            }
        }
    }

    private void addTextWatermark(Presentation presentation, PresentationAddWatermarkModel model) throws Exception {
        int width = 400;
        int height = 300;
        Rectangle2D.Double rect = new Rectangle2D.Double((presentation.getSlideSize().getSize().getWidth() - width) / 2,
                (presentation.getSlideSize().getSize().getHeight() - height) / 2, width, height);
        for (int i = 0; i < presentation.getSlides().getCount(); i++) {
            IAutoShape shape = presentation.getSlides().get(i).getShapes().appendShape(ShapeType.RECTANGLE, rect);
            shape.getFill().setFillType(FillFormatType.NONE);
            shape.getShapeStyle().getLineColor().setColor(Color.white);
            shape.setRotation(model.getRotate());
            shape.getLocking().setSelectionProtection(true);
            shape.getLine().setFillType(FillFormatType.NONE);
            shape.getTextFrame().setText(model.getText());
            PortionEx textRange = shape.getTextFrame().getTextRange();
            TextFont font = new TextFont(model.getFontName());
            textRange.setLatinFont(font);
            textRange.getFill().setFillType(FillFormatType.SOLID);
            textRange.getFill().getSolidColor().setColor(CommonHelper.String2Color(model.getColorText()));
            textRange.setFontHeight(model.getFontSize());
        }
    }

    private void addImageWatermark(Presentation presentation, byte[] imageData) throws Exception {
        ByteArrayInputStream bais = new ByteArrayInputStream(imageData);
        BufferedImage bufferedImage = ImageIO.read(bais);
        IImageData image = presentation.getImages().append(bufferedImage);
        for (int i = 0; i < presentation.getSlides().getCount(); i++) {
            ISlide slide = presentation.getSlides().get(i);
            slide.getSlideBackground().setType(BackgroundType.CUSTOM);
            slide.getSlideBackground().getFill().setFillType(FillFormatType.PICTURE);
            slide.getSlideBackground().getFill().getPictureFill().setFillType(PictureFillType.STRETCH);
            slide.getSlideBackground().getFill().getPictureFill().getPicture().setEmbedImage(image);
        }
    }

    private ChartType getChartType(String chartName) {
        ChartType chartType = ChartType.COLUMN_CLUSTERED;
        switch (chartName) {
            case "ColumnClustered":
                chartType = ChartType.COLUMN_CLUSTERED;
                break;
            case "ColumnStacked":
                chartType = ChartType.COLUMN_STACKED;
                break;
            case "Column3DClustered":
                chartType = ChartType.COLUMN_3_D_CLUSTERED;
                break;
            case "Column3DStacked":
                chartType = ChartType.COLUMN_3_D_STACKED;
                break;
            case "Column3D":
                chartType = ChartType.COLUMN_3_D;
                break;
            case "BarClustered":
                chartType = ChartType.BAR_CLUSTERED;
                break;
            case "BarStacked":
                chartType = ChartType.BAR_STACKED;
                break;
            case "Bar3DClustered":
                chartType = ChartType.BAR_3_D_CLUSTERED;
                break;
            case "Bar3DStacked":
                chartType = ChartType.BAR_3_D_STACKED;
                break;
            case "Line":
                chartType = ChartType.LINE;
                break;
            case "LineStacked":
                chartType = ChartType.LINE_STACKED;
                break;
            case "LineMarkers":
                chartType = ChartType.LINE_MARKERS;
                break;
            case "LineMarkersStacked":
                chartType = ChartType.LINE_MARKERS_STACKED;
                break;
            case "Line3D":
                chartType = ChartType.LINE_3_D;
                break;
            case "Pie":
                chartType = ChartType.PIE;
                break;
            case "Pie3D":
                chartType = ChartType.PIE_3_D;
                break;
            case "PieOfPie":
                chartType = ChartType.PIE_OF_PIE;
                break;
            case "PieExploded":
                chartType = ChartType.PIE_EXPLODED;
                break;
            case "Pie3DExploded":
                chartType = ChartType.PIE_3_D_EXPLODED;
                break;
            case "PieBar":
                chartType = ChartType.PIE_BAR;
                break;
            case "ScatterMarkers":
                chartType = ChartType.SCATTER_MARKERS;
                break;
            case "ScatterSmoothedLineMarkers":
                chartType = ChartType.SCATTER_SMOOTH_LINES_AND_MARKERS;
                break;
            case "ScatterSmoothedLine":
                chartType = ChartType.SCATTER_SMOOTH_LINES;
                break;
            case "ScatterLineMarkers":
                chartType = ChartType.SCATTER_STRAIGHT_LINES_AND_MARKERS;
                break;
            case "ScatterLine":
                chartType = ChartType.SCATTER_STRAIGHT_LINES;
                break;
            case "Area":
                chartType = ChartType.AREA;
                break;
            case "AreaStacked":
                chartType = ChartType.AREA_STACKED;
                break;
            case "Area3D":
                chartType = ChartType.AREA_3_D;
                break;
            case "Area3DStacked":
                chartType = ChartType.AREA_3_D_STACKED;
                break;
            case "Doughnut":
                chartType = ChartType.DOUGHNUT;
                break;
            case "DoughnutExploded":
                chartType = ChartType.DOUGHNUT_EXPLODED;
                break;
            case "Radar":
                chartType = ChartType.RADAR;
                break;
            case "RadarMarkers":
                chartType = ChartType.RADAR_MARKERS;
                break;
            case "RadarFilled":
                chartType = ChartType.RADAR_FILLED;
                break;
            case "Surface3D":
                chartType = ChartType.SURFACE_3_D;
                break;
            case "Surface3DNoColor":
                chartType = ChartType.SURFACE_3_D_NO_COLOR;
                break;
            case "Bubble":
                chartType = ChartType.BUBBLE;
                break;
            case "Bubble3D":
                chartType = ChartType.BUBBLE_3_D;
                break;
            case "CylinderClustered":
                chartType = ChartType.CYLINDER_CLUSTERED;
                break;
            case "CylinderStacked":
                chartType = ChartType.CYLINDER_STACKED;
                break;
            case "Cylinder3DClustered":
                chartType = ChartType.CYLINDER_3_D_CLUSTERED;
                break;
            case "ConeClustered":
                chartType = ChartType.CONE_CLUSTERED;
                break;
            case "ConeStacked":
                chartType = ChartType.CONE_STACKED;
                break;
            case "Cone3DClustered":
                chartType = ChartType.CONE_3_D_CLUSTERED;
                break;
            case "PyramidClustered":
                chartType = ChartType.PYRAMID_CLUSTERED;
                break;
            case "PyramidStacked":
                chartType = ChartType.PYRAMID_STACKED;
                break;
            case "Pyramid3DClustered":
                chartType = ChartType.PYRAMID_3_D_CLUSTERED;
                break;
        }
        return chartType;
    }

    private DataTable getDataTable() throws Exception {
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add(new DataColumn("SalesPers", DataTypes.DATATABLE_STRING));
        dataTable.getColumns().add(new DataColumn("SaleAmt", DataTypes.DATATABLE_INT));
        dataTable.getColumns().add(new DataColumn("ComPct", DataTypes.DATATABLE_INT));
        dataTable.getColumns().add(new DataColumn("ComAmt", DataTypes.DATATABLE_INT));
        DataRow row1 = dataTable.newRow();
        row1.setString("SalesPers", "Joe");
        row1.setInt("SaleAmt", 250);
        row1.setInt("ComPct", 150);
        row1.setInt("ComAmt", 99);
        DataRow row2 = dataTable.newRow();
        row2.setString("SalesPers", "Robert");
        row2.setInt("SaleAmt", 270);
        row2.setInt("ComPct", 150);
        row2.setInt("ComAmt", 99);
        DataRow row3 = dataTable.newRow();
        row3.setString("SalesPers", "Michelle");
        row3.setInt("SaleAmt", 310);
        row3.setInt("ComPct", 120);
        row3.setInt("ComAmt", 49);
        DataRow row4 = dataTable.newRow();
        row4.setString("SalesPers", "Erich");
        row4.setInt("SaleAmt", 330);
        row4.setInt("ComPct", 120);
        row4.setInt("ComAmt", 49);
        DataRow row5 = dataTable.newRow();
        row5.setString("SalesPers", "Dafna");
        row5.setInt("SaleAmt", 360);
        row5.setInt("ComPct", 150);
        row5.setInt("ComAmt", 141);
        DataRow row6 = dataTable.newRow();
        row6.setString("SalesPers", "Rob");
        row6.setInt("SaleAmt", 380);
        row6.setInt("ComPct", 150);
        row6.setInt("ComAmt", 135);
        dataTable.getRows().add(row1);
        dataTable.getRows().add(row2);
        dataTable.getRows().add(row3);
        dataTable.getRows().add(row4);
        dataTable.getRows().add(row5);
        dataTable.getRows().add(row6);
        return dataTable;
    }

    private void insertDatatableToChart(IChart chart, DataTable dataTable) throws Exception {
        for (int c = 0; c < dataTable.getColumns().size(); c++) {
            chart.getChartData().get(0, c).setText(dataTable.getColumns().get(c).getColumnName());
        }
        for (int r = 0; r < dataTable.getRows().size(); r++) {
            Object[] datas = dataTable.getRows().get(r).getArrayList();
            for (int c = 0; c < datas.length; c++) {
                chart.getChartData().get(r + 1, c).setValue(datas[c]);
            }
        }
    }

}

