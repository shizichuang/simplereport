package com.example.simplereport.common;

import lombok.Data;

@Data
public class ExcelSelectorContainer {
    /**
     * 第几个sheet，从0开始
     */
    private int sheetIndex;
    /**
     * 下拉单元格行号，从0开始
     */
    private int firstRow;
    /**
     * 下拉单元格结束行号
     */
    private int lastRow;
    /**
     * 下拉单元格列号，从0开始
     */
    private int firstCol;
    /**
     * 下拉单元格列（结束）
     */
    private int lastCol;
    /**
     * 动态生成的下拉内容，easyPoi使用的是字符数组
     */
    private String[] datas;

    /**
     * 指定实体属性名
     */
    private String fieldName;

    /**
     * 是否遍历下拉
     */
    private boolean ForForeach;
}
