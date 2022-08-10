package com.iceblue.livedemo.model.powerpoint;

/**
 * 查找或替换文本的请求model
 */

public class PresentationFindAndReplaceModel {
    private String findText;
    private String newText;
    private String typeChance;

    public String getFindText() {
        return findText;
    }

    public void setFindText(String findText) {
        this.findText = findText;
    }

    public String getNewText() {
        return newText;
    }

    public void setNewText(String newText) {
        this.newText = newText;
    }

    public String getTypeChance() {
        return typeChance;
    }

    public void setTypeChance(String typeChance) {
        this.typeChance = typeChance;
    }
}
