package com.iceblue.livedemo.service;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.TypeOptionsModel;
import com.iceblue.livedemo.model.word.*;
import com.iceblue.livedemo.utils.CommonHelper;
import com.iceblue.livedemo.utils.Static;
import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TextRange;
import com.spire.doc.formatting.CellFormat;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.awt.geom.Dimension2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.*;
import java.util.List;

/**
 * 产品业务逻辑都在Service
 * 结果文件保存格式
 * 路径：Static.OUTPUT_FILE_PATH
 * 文件名：Static.OUTPUT_FILE_START+UUID.randomUUID()+".pdf" (后缀视情况改为对应的格式)
 */
@Service
public class WordDemoService {

    public void addWaterMark(ResultModel resultModel, Document doc, byte[] imgData, WordAddWaterMarkRequestModel model) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".docx";

        switch (model.getWaterMarkType()) {
            case "Text":
                addTextWatermark(doc, model);
                break;
            case "Image":
                addImageWatermark(doc, imgData);
                break;
            default:
                throw new RuntimeException();
        }
        doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.Docx);

        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }


    public void mailMerge(ResultModel resultModel, WordMailMergeRequestModel model) throws Exception {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".docx";
        Document doc = new Document();
        doc.loadFromFile(Static.WORD_TEMPLATES_FILE_PATH + "Template.doc");
        String[] values = {model.getAgreementStartDateStr(),
                model.getAgreementEndDateStr(),
                model.getAgreementExtensionDateStr(),
                model.getDocumentationStartDateStr(),
                model.getDocumentationEndDateStr()
        };
        String[] fields = {
                "SubGrantPAStartDateValue",
                "SubGrantPAEndDateValue",
                "SubGrantPAExtensionDateValue",
                "SubGrantPSStartDateValue",
                "SubGrantPSEndDateValue"
        };
        doc.getMailMerge().execute(fields, values);
        doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.Docx);
        doc.close();
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void addTable(ResultModel resultModel, WordTableRequestModel model) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".docx";
        //Get Colors
        Color headerBack = CommonHelper.String2Color(model.getHeaderBackColor());
        Color rowBack = CommonHelper.String2Color(model.getRowBackColor());
        Color alternationBack = CommonHelper.String2Color(model.getAlternationBackColor());
        Color borderBack = CommonHelper.String2Color(model.getBorderColor());
        //Get table
        List<CountryModel> dataList = CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "wordTable.xml", CountryModel.class);

        Document doc = new Document();
        doc.loadFromFile(Static.WORD_TEMPLATES_FILE_PATH + "Blank.doc");
        Section section = doc.getSections().get(0);
        Table table = section.addTable();
        int rowCount = dataList.size();
        int columnCount = CountryModel.class.getDeclaredFields().length;
        table.setDefaultRowHeight(25f);
        table.setDefaultColumnWidth(0f);
        table.resetCells(rowCount + 1, columnCount);
        //Borders
        Borders borders = table.getTableFormat().getBorders();
        //Left
        borders.getLeft().setBorderType(BorderStyle.Hairline);
        borders.getLeft().setColor(borderBack);
        //Right
        borders.getRight().setBorderType(BorderStyle.Hairline);
        borders.getRight().setColor(borderBack);
        //Bottom
        borders.getBottom().setBorderType(BorderStyle.Hairline);
        borders.getBottom().setColor(borderBack);
        //Top
        borders.getTop().setBorderType(BorderStyle.Hairline);
        borders.getTop().setColor(borderBack);
        //Horizontal
        borders.getHorizontal().setBorderType(BorderStyle.Hairline);
        borders.getHorizontal().setColor(borderBack);
        //Vertical
        borders.getVertical().setBorderType(BorderStyle.Hairline);
        borders.getVertical().setColor(borderBack);

        //Table header
        TableRow headerRow = table.getRows().get(0);
        headerRow.isHeader(true);
        for (int i = 0; i < columnCount; i++) {
            Paragraph p = headerRow.getCells().get(i).addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange headerText = p.appendText(CommonHelper.lowerOrUpper(CountryModel.class.getDeclaredFields()[i].getName()));
            headerText.getCharacterFormat().setBold(true);
            CellFormat cellStyle = headerRow.getCells().get(i).getCellFormat();
            cellStyle.setVerticalAlignment(VerticalAlignment.Middle);
            cellStyle.setBackColor(headerBack);
        }
        //Data row
        for (int i = 0; i < rowCount; i++) {
            CountryModel countryModel = dataList.get(i);
            for (int j = 0; j < columnCount; j++) {
                Paragraph p = table.getRows().get(i + 1).getCells().get(j).addParagraph();
                //获取对象的属性名
                String fieldName = CountryModel.class.getDeclaredFields()[j].getName();
                //拼凑get方法，
                Object object = CommonHelper.Reflection(countryModel, CountryModel.class, "get" + CommonHelper.lowerOrUpper(fieldName), null);
                if (!fieldName.equals("flag")) {
                    p.appendText(String.valueOf(object));
                } else {
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                    byte[] imgData = Base64.getDecoder().decode(String.valueOf(object));
                    p.appendPicture(imgData);
                }
                CellFormat cellFormat = table.getRows().get(i + 1).getCells().get(j).getCellFormat();
                cellFormat.setVerticalAlignment(VerticalAlignment.Middle);
                cellFormat.setBackColor(rowBack);
                if (alternationBack != null && i % 2 == 1) {
                    cellFormat.setBackColor(alternationBack);
                }
            }
        }

        doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.Docx);
        doc.close();
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);

    }

    public void addHeaderAndFooter(ResultModel resultModel, Document doc, byte[] headerImgData, byte[] footerImgData, WordHeaderAndFooterRequestModel model) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".docx";
        for (int i = 0; i < doc.getSections().getCount(); i++) {
            //Page size
            Section section = doc.getSections().get(i);
            section.getPageSetup().setPageSize(getPageSize(model.getPageSize()));
            section.getPageSetup().getMargins().setTop(72f);
            section.getPageSetup().getMargins().setBottom(72f);
            section.getPageSetup().getMargins().setLeft(89f);
            section.getPageSetup().getMargins().setRight(89f);
            //Header and footer
            HeaderFooter header = section.getHeadersFooters().getHeader();
            HeaderFooter footer = section.getHeadersFooters().getFooter();
            //Header
            //Header picture
            Paragraph headerParagraph = header.addParagraph();
            DocPicture headerPicture = headerParagraph.appendPicture(headerImgData);
            headerPicture.setTextWrappingStyle(TextWrappingStyle.Behind);
            headerPicture.setHorizontalOrigin(HorizontalOrigin.Page);
            headerPicture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
            headerPicture.setVerticalOrigin(VerticalOrigin.Page);
            headerPicture.setVerticalAlignment(ShapeVerticalAlignment.Top);
            //Header text
            TextRange text = headerParagraph.appendText(model.getHeaderText());
            text.getCharacterFormat().setFontName(model.getHeaderFont());
            text.getCharacterFormat().setFontSize(model.getHeaderFontSize());
            text.getCharacterFormat().setTextColor(CommonHelper.String2Color(model.getHeaderTextColor()));
            headerParagraph.getFormat().setHorizontalAlignment(getHorizontalAlignment(model.getHeaderTextAlignment()));
            //border
            headerParagraph.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Hairline);
            headerParagraph.getFormat().getBorders().getBottom().setSpace(0.03f);

            //Footer text
            Paragraph footerParagraph = footer.addParagraph();
            text = footerParagraph.appendText(model.getFooterText());
            text.getCharacterFormat().setFontName(model.getFooterFont());
            text.getCharacterFormat().setFontSize(model.getFooterFontSize());
            text.getCharacterFormat().setTextColor(CommonHelper.String2Color(model.getFooterTextColor()));
            footerParagraph.getFormat().setHorizontalAlignment(getHorizontalAlignment(model.getFooterTextAlignment()));
            //border
            footerParagraph.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Hairline);
            footerParagraph.getFormat().getBorders().getBottom().setSpace(0.03f);
            //Footer
            DocPicture footerPicture = footerParagraph.appendPicture(footerImgData);
            footerPicture.setTextWrappingStyle(TextWrappingStyle.Behind);
            footerPicture.setHorizontalOrigin(HorizontalOrigin.Page);
            footerPicture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
            footerPicture.setVerticalOrigin(VerticalOrigin.Page);
            footerPicture.setVerticalAlignment(ShapeVerticalAlignment.Bottom);

            //Insert page number
            footerParagraph = footer.addParagraph();
            footerParagraph.appendField("page number", FieldType.Field_Page);
            footerParagraph.appendText("of");
            footerParagraph.appendField("number of pages", FieldType.Field_Num_Pages);
            footerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
        }
        doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.Docx);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void wordConversion(ResultModel resultModel, Document doc, String convertTo) throws IOException {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID();
        switch (convertTo) {
            case "PDF":
                outputFile += ".pdf";
                doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.PDF);
                break;
            case "XPS":
                outputFile += ".xps";
                doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.XPS);
                break;
            case "IMAGE":
                BufferedImage[] images = doc.saveToImages(ImageType.Bitmap);
                outputFile = CommonHelper.packageImages(images, "png");
                break;
            case "RTF":
                outputFile += ".rtf";
                doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.Rtf);
                break;
            case "TIFF":
                outputFile += ".tiff";
                doc.saveToTiff(Static.OUTPUT_FILE_PATH + outputFile);
                break;
            case "HTML":
                outputFile = packageHtml(doc);
                break;
        }
        doc.close();
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void findAndHighlight(ResultModel resultModel, Document doc, String text) {
        String outputFile = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".docx";
        TextSelection[] textSelections = doc.findAllString(text, false, true);
        if (textSelections != null) {
            for (TextSelection selection : textSelections) {
                selection.getAsOneRange().getCharacterFormat().setHighlightColor(Color.YELLOW);
            }
        }
        doc.saveToFile(Static.OUTPUT_FILE_PATH + outputFile, FileFormat.Docx);
        doc.close();
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(outputFile);
    }

    public void getConvertTypeOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "wordConvertOptions.xml", TypeOptionsModel.class));
    }

    public void getPageSizeOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "wordPageSizeOptions.xml", TypeOptionsModel.class));
    }

    public void getTextAlignmentOptions(ResultModel resultModel) {
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "wordTextAlignmentOptions.xml", TypeOptionsModel.class));
    }

    public void getTableData(ResultModel resultModel) {
        List<CountryModel> dataList = CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "wordTable.xml", CountryModel.class);
        resultModel.setValid(true);
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setData(dataList);
    }

    /******************************************************************************************/
    private void addImageWatermark(Document doc, byte[] imgData) {
        PictureWatermark pictureWatermark = new PictureWatermark();
        pictureWatermark.setPicture(new ByteArrayInputStream(imgData));
        pictureWatermark.setScaling(250);
        pictureWatermark.isWashout(false);
        doc.setWatermark(pictureWatermark);
    }

    private void addTextWatermark(Document doc, WordAddWaterMarkRequestModel model) {
        TextWatermark textWatermark = new TextWatermark();
        textWatermark.setText(model.getWaterMarkText());
        textWatermark.setFontSize(model.getTextSize());
        textWatermark.setColor(CommonHelper.String2Color(model.getTextColor()));
        textWatermark.setLayout(WatermarkLayout.Diagonal);
        doc.setWatermark(textWatermark);
    }

    private Dimension2D getPageSize(String pageSizeName) {
        Dimension2D pageSize = null;
        switch (pageSizeName) {
            case "A3":
                pageSize = PageSize.A3;
                break;
            case "A4":
                pageSize = PageSize.A4;
                break;
            case "A5":
                pageSize = PageSize.A5;
                break;
            case "A6":
                pageSize = PageSize.A6;
                break;
            case "B4":
                pageSize = PageSize.B4;
                break;
            case "B5":
                pageSize = PageSize.B5;
                break;
            case "B6":
                pageSize = PageSize.B6;
                break;
            case "Letter":
                pageSize = PageSize.Letter;
                break;
            case "Letter11x17":
                pageSize = PageSize.Letter_11_x_17;
                break;
            case "EnvelopeDL":
                pageSize = PageSize.Envelope_DL;
                break;
            case "Quarto":
                pageSize = PageSize.Quarto;
                break;
            case "Statement":
                pageSize = PageSize.Statement;
                break;
            case "Ledger":
                pageSize = PageSize.Ledger;
                break;
            case "Tabloid":
                pageSize = PageSize.Tabloid;
                break;
            case "Note":
                pageSize = PageSize.Note;
                break;
            case "Legal":
                pageSize = PageSize.Legal;
                break;
            case "Flsa":
                pageSize = PageSize.Flsa;
                break;
            case "Executive":
                pageSize = PageSize.Executive;
                break;
        }
        return pageSize;
    }

    private HorizontalAlignment getHorizontalAlignment(String alignment) {
        HorizontalAlignment horizontalAlignment = null;
        switch (alignment) {
            case "Left":
                horizontalAlignment = HorizontalAlignment.Left;
                break;
            case "Right":
                horizontalAlignment = HorizontalAlignment.Right;
                break;
            case "Center":
                horizontalAlignment = HorizontalAlignment.Center;
                break;
            case "Distribute":
                horizontalAlignment = HorizontalAlignment.Distribute;
                break;
            case "HightKashida":
                horizontalAlignment = HorizontalAlignment.Hight_Kashida;
                break;
            case "Justify":
                horizontalAlignment = HorizontalAlignment.Justify;
                break;
            case "LowKashida":
                horizontalAlignment = HorizontalAlignment.Low_Kashida;
                break;
            case "MediumKashida":
                horizontalAlignment = HorizontalAlignment.Medium_Kashida;
                break;
        }
        return horizontalAlignment;
    }

    private String packageHtml(Document doc) throws IOException {
        //随机串 用于保证每个请求得到的folder都不一样 避免冲突
        String ID = UUID.randomUUID().toString();
        //html名称
        String htmlFile = ID + ".html";

        doc.getHtmlExportOptions().setImageEmbedded(true);
        doc.getHtmlExportOptions().setCssStyleSheetType(CssStyleSheetType.Internal);
        //转html
        doc.saveToFile(Static.OUTPUT_FILE_PATH + htmlFile, FileFormat.Html);

        //返回zip文件名
        return htmlFile;
    }

}

