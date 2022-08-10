package com.iceblue.livedemo.utils;

import com.iceblue.livedemo.model.ResultModel;
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.pdf.PdfDocument;
import com.spire.presentation.Presentation;
import com.spire.xls.ExcelVersion;
import com.spire.xls.Workbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;


/**
 * 用于初始化输入文档
 */
public class FileInitHelper {

    public static byte[] initImage(ResultModel resultModel,MultipartFile imgFile,String template){
        byte[] data;
        FileInputStream fis=null;
        if(imgFile==null||imgFile.isEmpty()){
            try {
                fis=new FileInputStream(Static.IMAGE_TEMPLATE_PATH+template);
                data=new byte[fis.available()];
                fis.read(data,0,data.length);
                return data;
            }catch (IOException e){
                resultModel.setMessage(Static.MESSAGE_OTHER_ERROR);
                resultModel.setValid(false);
                return null;
            }finally {
                if(fis!=null){
                    try {
                        fis.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        try {
            return imgFile.getBytes();
        } catch (IOException e) {
            e.printStackTrace();
            resultModel.setMessage(Static.MESSAGE_FAILED_PARSE_IMAGE);
            resultModel.setValid(false);
            return null;
        }
    }

    public static Document initWord(ResultModel resultModel, MultipartFile docFile,String template){
        Document doc=null;
        //如果上传的文档为空 那么默认使用模板文档
        if(docFile==null||docFile.isEmpty()){
            doc=new Document();
            doc.loadFromFile(Static.WORD_TEMPLATES_FILE_PATH +template);
        }else{
            //否则判断文档格式并加载上传的文档
            try {
                String originalFileName = docFile.getOriginalFilename();
                String exName = originalFileName.substring(originalFileName.lastIndexOf(".") + 1).toLowerCase();

                com.spire.doc.FileFormat format=null;
                switch (exName){
                    case "docx":
                        format= FileFormat.Docx;
                        break;
                    case "doc":
                        format= FileFormat.Doc;
                        break;
                    case "txt":
                        format= FileFormat.Txt;
                        break;
                    case "rtf":
                        format= FileFormat.Rtf;
                        break;
                    default:
                        resultModel.setValid(false);
                        resultModel.setMessage(Static.MESSAGE_WRONG_FORMAT);
                        return null;
                }
                doc=new Document();
                ByteArrayInputStream inputStream=new ByteArrayInputStream(docFile.getBytes());
                doc.loadFromStream(inputStream,format);
            }catch (Exception e){
                e.printStackTrace();
                resultModel.setValid(false);
                resultModel.setMessage(Static.MESSAGE_FAILED_PARSE_FILE);
                //出异常则返回null
                return null;
            }
        }

        return doc;
    }


    public static PdfDocument initPDF(ResultModel resultModel, MultipartFile docFile,String template){
        PdfDocument pdf=null;
        if(docFile==null||docFile.isEmpty()) {
            pdf = new PdfDocument();
            pdf.loadFromFile(Static.PDF_TEMPLATES_FILE_PATH + template);
        }else{
            try {
                String originalFileName = docFile.getOriginalFilename();
                String exName = originalFileName.substring(originalFileName.lastIndexOf(".") + 1).toLowerCase();
                //检查文件格式是否正确
                if(!(exName.equals("pdf"))){
                    resultModel.setValid(false);
                    resultModel.setMessage(Static.MESSAGE_WRONG_FORMAT);
                    return null;
                }
                pdf=new PdfDocument();
                ByteArrayInputStream inputStream=new ByteArrayInputStream(docFile.getBytes());

                pdf.loadFromStream(inputStream);
            }catch (Exception e){
                e.printStackTrace();
                resultModel.setValid(false);
                resultModel.setMessage(Static.MESSAGE_FAILED_PARSE_FILE);
                //出异常则返回null
                return null;
            }
        }
        return pdf;
    }

    public static Workbook initExcel(ResultModel resultModel, MultipartFile excelFile,String template){
        Workbook workbook=null;
        //如果上传的文档为空 那么默认使用模板文档
        if(excelFile ==null||excelFile.isEmpty()){
            workbook=new Workbook();
            workbook.loadFromFile(Static.Excel_TEMPLATES_FILE_PATH +template);
        }else{
            //否则判断文档格式并加载上传的文档
            try {
                String originalFileName = excelFile.getOriginalFilename();
                String exName = originalFileName.substring(originalFileName.lastIndexOf(".") + 1).toLowerCase();
                com.spire.xls.ExcelVersion excelVersion=null;
                //检查文件格式是否正确
                switch (exName){
                    case "xlsx":
                        excelVersion= ExcelVersion.Version2016;
                        break;
                    case "xls":
                        excelVersion=ExcelVersion.Version97to2003;
                        break;
                    case "xlsb":
                        excelVersion=ExcelVersion.Xlsb2010;
                        break;
                    case "ods":
                        excelVersion=ExcelVersion.ODS;
                        break;
                    default:
                        resultModel.setValid(false);
                        resultModel.setMessage(Static.MESSAGE_WRONG_FORMAT);
                        return null;
                }

                workbook=new Workbook();
                ByteArrayInputStream inputStream=new ByteArrayInputStream(excelFile.getBytes());

                workbook.loadFromStream(inputStream,excelVersion);
            }catch (Exception e){
                e.printStackTrace();
                resultModel.setValid(false);
                resultModel.setMessage(Static.MESSAGE_FAILED_PARSE_FILE);
                //出异常则返回null
                return null;
            }
        }
        return workbook;
    }

    /**
     * 初始化PPT文档
     * @param pptFile
     * @param template
     * @return
     * @throws Exception
     */
    public static Presentation initPPT(ResultModel resultModel, MultipartFile pptFile,String template){
        Presentation presentation=null;
        //如果上传的文档为空 那么默认使用模板文档
        if(pptFile ==null||pptFile.isEmpty()){
            presentation = new Presentation();
            try{
                presentation.loadFromFile(Static.Ppt_TEMPLATES_FILE_PATH +template);
            }catch (Exception e){
                e.printStackTrace();
                resultModel.setValid(false);
                resultModel.setMessage(Static.MESSAGE_FAILED_PARSE_FILE);
                return null;
            }
        }else{
            //否则判断文档格式并加载上传的文档
            try {
                String originalFileName = pptFile.getOriginalFilename();
                String exName = originalFileName.substring(originalFileName.lastIndexOf(".") + 1).toLowerCase();
                //检查文件格式是否正确
                if(!(exName.equals("pptx")||exName.equals("ppt"))){
                    resultModel.setValid(false);
                    resultModel.setMessage(Static.MESSAGE_WRONG_FORMAT);
                    return null;
                }
                presentation = new Presentation();
                ByteArrayInputStream inputStream=new ByteArrayInputStream(pptFile.getBytes());
                com.spire.presentation.FileFormat format="pptx".equals(exName)? com.spire.presentation.FileFormat.PPTX_2016 : com.spire.presentation.FileFormat.PPT;
                presentation.loadFromStream(inputStream,format);
            }catch (Exception e){
                e.printStackTrace();
                resultModel.setValid(false);
                resultModel.setMessage(Static.MESSAGE_FAILED_PARSE_FILE);
                //出异常则返回null
                return null;
            }
        }
        return presentation;
    }
}
