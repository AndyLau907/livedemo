package com.iceblue.livedemo.controller;

import com.iceblue.livedemo.model.ResultModel;
import com.iceblue.livedemo.model.TypeOptionsModel;
import com.iceblue.livedemo.utils.CommonHelper;
import com.iceblue.livedemo.utils.Static;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@RequestMapping(value = "/common")
@RestController
public class CommonController {
    @RequestMapping(value = "/downloadOutputFile")
    public void download(@RequestParam(value = "fileName") String fileName) throws IOException {
        ServletRequestAttributes requestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = requestAttributes.getResponse();
        CommonHelper.downloadFile(response, Static.OUTPUT_FILE_PATH, fileName, true);
    }

    @RequestMapping(method = RequestMethod.GET, value = "/getFontOptions")
    public ResultModel getFontOptions() {
        ResultModel resultModel = new ResultModel();
        resultModel.setMessage(Static.MESSAGE_SUCCESS);
        resultModel.setValid(true);
        resultModel.setData(CommonHelper.xmlToBean(Static.DATA_FILE_PATH + "fontOptions.xml", TypeOptionsModel.class));
        return resultModel;
    }
}
