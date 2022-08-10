package com.iceblue.livedemo.service;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.TypeOptionsModel;
import com.iceblue.livedemo.model.pdf.CustomerModel;
import com.iceblue.livedemo.model.pdf.PdfAddWaterMarkRequestModel;
import com.iceblue.livedemo.model.pdf.PdfFindAndHighlightRequestModel;
import com.iceblue.livedemo.model.pdf.PdfTableRequestModel;
import com.iceblue.livedemo.model.word.CountryModel;
import com.iceblue.livedemo.utils.CommonHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.pdf.*;
import com.spire.pdf.conversion.PdfStandardsConverter;
import com.spire.pdf.general.find.PdfTextFind;
import com.spire.pdf.general.find.TextFindParameter;
import com.spire.pdf.graphics.*;
import com.spire.pdf.grid.PdfGrid;
import com.spire.pdf.grid.PdfGridCell;
import com.spire.pdf.grid.PdfGridRow;
import org.apache.commons.lang.StringUtils;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Dimension2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.EnumSet;
import java.util.List;
import java.util.UUID;

@Service
public class PdfDemoService {


    public void pdfConversion(ResultModel resultModel, PdfDocument pdf, MultipartFile pdfFile, String convertType) throws IOException {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID();
        PdfStandardsConverter standardsConverter = null;
        switch (convertType) {
            case "DOCX":
                outputFile += ".docx";
                pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.DOCX);
                break;
            case "DOC":
                outputFile += ".doc";
                pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.DOC);
                break;
            case "XPS":
                outputFile += ".xps";
                pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.XPS);
                break;
            case "XLSX":
                outputFile += ".xlsx";
                pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.XLSX);
                break;
            case "POSTSCRIPT":
                outputFile += ".ps";
                pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.POSTSCRIPT);
                break;
            case "HTML":
                outputFile += ".html";
                pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.HTML);
                break;
            case "IMAGE":
                outputFile = CommonHelper.packageImages(saveToImages(pdf), "png");
                break;
            case "SVG":
                outputFile = packageSvgs(pdf);
                break;
            case "TIFF":
                outputFile += ".tiff";
                pdf.saveToTiff(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "PDFA1A":
                outputFile += ".pdf";
                standardsConverter = new PdfStandardsConverter(new ByteArrayInputStream(pdfFile.getBytes()));
                standardsConverter.toPdfA1A(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "PDFA1B":
                outputFile += ".pdf";
                standardsConverter = new PdfStandardsConverter(new ByteArrayInputStream(pdfFile.getBytes()));
                standardsConverter.toPdfA1B(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "PDFA2A":
                outputFile += ".pdf";
                standardsConverter = new PdfStandardsConverter(new ByteArrayInputStream(pdfFile.getBytes()));
                standardsConverter.toPdfA2A(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "PDFA2B":
                outputFile += ".pdf";
                standardsConverter = new PdfStandardsConverter(new ByteArrayInputStream(pdfFile.getBytes()));
                standardsConverter.toPdfA2B(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "PDFA3A":
                outputFile += ".pdf";
                standardsConverter = new PdfStandardsConverter(new ByteArrayInputStream(pdfFile.getBytes()));
                standardsConverter.toPdfA3A(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "PDFA3B":
                outputFile += ".pdf";
                standardsConverter = new PdfStandardsConverter(new ByteArrayInputStream(pdfFile.getBytes()));
                standardsConverter.toPdfA3B(Static.OUTPUT_FILE_PATH + outputFile);
                break;
        }
        pdf.close();
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void drawTable(ResultModel resultModel, PdfTableRequestModel model) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pdf";
        PdfDocument pdf = new PdfDocument();
        PdfPageBase page = pdf.getPages().add(PdfPageSize.A4, new PdfMargins(40, 30, 40, 30));
        PdfBorders borders = new PdfBorders();
        if (StringUtils.isBlank(model.getBorderColor()) || CommonHelper.String2Color(model.getBorderColor()) == null) {
            //如果没传颜色值或者颜色值错误 默认为透明边框 即没有边框
            borders.setAll(new PdfPen(new PdfRGBColor(Color.TRANSLUCENT)));
        } else {
            borders.setAll(new PdfPen(new PdfRGBColor(CommonHelper.String2Color(model.getBorderColor()))));
        }

        PdfGrid grid = null;
        PdfGridRow pdfGridRow;
        PdfTrueTypeFont font1 = new PdfTrueTypeFont(new Font("Arial", Font.BOLD, 16), true);
        PdfTrueTypeFont font2 = new PdfTrueTypeFont(new Font("Arial", Font.PLAIN, 10), true);
        PdfStringFormat format1 = new PdfStringFormat(PdfTextAlignment.Center);
        grid = new PdfGrid();
        grid.setAllowCrossPages(true);
        grid.getStyle().setCellSpacing(3f);
        grid.getColumns().add(5);

        double width = page.getCanvas().getClientSize().getWidth() - (grid.getColumns().getCount() + 1);
        grid.getColumns().get(0).setWidth(width * 0.2);
        grid.getColumns().get(1).setWidth(width * 0.2);
        grid.getColumns().get(2).setWidth(width * 0.2);
        grid.getColumns().get(3).setWidth(width * 0.2);
        grid.getColumns().get(4).setWidth(width * 0.2);
        pdfGridRow = grid.getHeaders().add(1)[0];
        pdfGridRow.setHeight(50);
        int i = 0;
        List<CustomerModel> data = CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pdfTable.xml", CustomerModel.class);
        Field[] fields = CustomerModel.class.getDeclaredFields();
        //header row
        for (i = 0; i < fields.length; i++) {
            PdfGridCell cell = pdfGridRow.getCells().get(i);

            cell.setColumnSpan(1);
            cell.setValue(CommonHelper.lowerOrUpper(fields[i].getName()));
            cell.setStringFormat(format1);
            cell.getStyle().setBorders(borders);
            cell.getStyle().setFont(font1);
        }
        //data row
        for (i = 0; i < data.size(); i++) {
            pdfGridRow = grid.getRows().add();
            pdfGridRow.setHeight(40f);
            for (int j = 0; j < fields.length; j++) {
                PdfGridCell cell = pdfGridRow.getCells().get(j);
                cell.setColumnSpan(1);
                cell.setValue(CommonHelper.Reflection(data.get(i), data.get(i).getClass(), "get" + CommonHelper.lowerOrUpper(fields[j].getName()), null));
                cell.getStyle().setBorders(borders);
                cell.getStyle().setFont(font2);
                cell.setStringFormat(format1);
            }
        }
        grid.setRepeatHeader(model.isRepeatHeader());
        grid.draw(page, new Point(0, 0));
        pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PDF);

        pdf.close();
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void mergePdf(ResultModel resultModel, MultipartFile[] pdfs) throws IOException {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pdf";
        ByteArrayInputStream[] streams = new ByteArrayInputStream[pdfs.length];

        for (int i = 0; i < pdfs.length; i++) {
            streams[i] = new ByteArrayInputStream(pdfs[i].getBytes());
        }
        PdfDocumentBase documentBase = PdfDocument.mergeFiles(streams);
        documentBase.save(Static.OUTPUT_FILE_PATH + outputFile);
        documentBase.close();

        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void addWaterMark(ResultModel resultModel, PdfDocument pdf, byte[] imgData, PdfAddWaterMarkRequestModel model) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pdf";
        switch (model.getWaterMarkType()) {
            case "Text":
                addTextWatermark(pdf, model);
                break;
            case "Image":
                addImageWatermark(pdf, imgData);
                break;
            default:
                throw new RuntimeException();
        }
        pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PDF);
        pdf.close();

        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void findAndHighlight(ResultModel resultModel, PdfDocument pdf, PdfFindAndHighlightRequestModel model) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".pdf";
        PdfTextFind[] result = null;
        for (Object pageObj : pdf.getPages()) {
            PdfPageBase page = (PdfPageBase) pageObj;
            // Find text
            result = page.findText(model.getText(), EnumSet.of(TextFindParameter.None)).getFinds();
            if (result != null) {
                for (PdfTextFind find : result) {
                    find.applyHighLight(CommonHelper.String2Color(model.getColor()));
                }
            }
        }
        pdf.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PDF);
        pdf.close();

        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);

    }

    public void getConvertOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pdfConvertOptions.xml", TypeOptionsModel.class));
    }


    /*******************************************************************************************************/
    private String packageSvgs(PdfDocument pdf) throws IOException {
        String ID = UUID.randomUUID().toString();
        String outputFile = Static.OUTPUT_FILE_START + ID + ".zip";
        String svgName = "output.svg";
        String directoryPath = Static.OUTPUT_FILE_PATH + ID;
        //创建文件夹
        File htmlFolder = new File(directoryPath);
        if (!htmlFolder.exists()) {
            htmlFolder.mkdir();
        }
        pdf.saveToFile(directoryPath + "/" + svgName, FileFormat.SVG);
        FileOutputStream fos = new FileOutputStream(Static.OUTPUT_FILE_PATH + outputFile);
        CommonHelper.toZip(Static.OUTPUT_FILE_PATH + ID, fos, true);
        CommonHelper.deleteDirectory(directoryPath);
        return outputFile;
    }

    private BufferedImage[] saveToImages(PdfDocument pdf) {
        BufferedImage[] images = new BufferedImage[pdf.getPages().getCount()];
        for (int i = 0; i < pdf.getPages().getCount(); i++) {
            images[i] = pdf.saveAsImage(i, PdfImageType.Bitmap);
        }
        return images;
    }

    private void addImageWatermark(PdfDocument pdf, byte[] imgData) {
        PdfImage image = PdfImage.fromStream(new ByteArrayInputStream(imgData));
        for (int i = 0; i < pdf.getPages().getCount(); i++) {
            PdfPageBase page = pdf.getPages().get(i);
            page.getCanvas().save();
            page.getCanvas().setTransparency(0.5f, 0.5f, PdfBlendMode.Multiply);
            page.getCanvas().drawImage(image, new Point2D.Float(160, 260));
            page.getCanvas().restore();
        }
    }

    private void addTextWatermark(PdfDocument pdf, PdfAddWaterMarkRequestModel model) {
        for (int i = 0; i < pdf.getPages().getCount(); i++) {
            PdfPageBase page = pdf.getPages().get(i);
            Dimension2D dimension2D = new Dimension();
            dimension2D.setSize(page.getCanvas().getClientSize().getWidth() / 2, page.getCanvas().getClientSize().getHeight() / 3);
            PdfTilingBrush brush = new PdfTilingBrush(dimension2D);
            brush.getGraphics().setTransparency(0.3F);
            brush.getGraphics().save();
            brush.getGraphics().translateTransform((float) brush.getSize().getWidth() / 2, (float) brush.getSize().getHeight() / 2);
            brush.getGraphics().rotateTransform(-45);
            brush.getGraphics().drawString(model.getWaterMarkText(), new PdfTrueTypeFont(new Font(model.getTextFont(), Font.BOLD, model.getTextSize()), true), PdfBrushes.getViolet(), 0, 0, new PdfStringFormat(PdfTextAlignment.Center));
            brush.getGraphics().restore();
            brush.getGraphics().setTransparency(0.5);
            Rectangle2D loRect = new Rectangle2D.Float();
            loRect.setFrame(new Point2D.Float(0, 0), page.getCanvas().getClientSize());
            page.getCanvas().drawRectangle(brush, loRect);
        }
    }

    public void getTableData(ResultModel resultModel) {
        List<CustomerModel> dataList = CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "pdfTable.xml", CustomerModel.class);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(dataList);
    }
}
