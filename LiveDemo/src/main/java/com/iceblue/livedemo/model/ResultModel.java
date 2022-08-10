package com.iceblue.livedemo.model;

/**
 * 返回数据统一模型
 */
public class ResultModel {
    /**
     * 表示本次请求对应的后台操作是否执行成功，如果为false则代表执行失败
     * 失败可能是因为执行产品功能出现异常或错误了，或者客户没有填必填项的参数？
     * 必填项和参数合规性的验证应该可以放前端也可以后端验证
     */
    private boolean isValid;
    /**
     * 向前端返回的信息，比如出现异常时对应的提示等
     */
    private String message;
    /**
     * 所有所需要的数据都在data域。具体模型应该根据每个接口的需求不同再定
     */
    private Object data;

    public boolean isValid() {
        return isValid;
    }

    public void setValid(boolean valid) {
        isValid = valid;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }
}
