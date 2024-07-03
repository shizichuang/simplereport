package com.example.simplereport.common;

import lombok.Data;

import java.util.Map;

/**
 * 级联 容器 -- 后续建议依赖此容器搞一个公共的级联工具
 */
@Data
public class MultiLevelDropDownContainer {
    /**
     * 存放值的sheet
     */
    private String sheetName;
    /**
     * 下拉单元格行号，从0开始
     */
    private int firstRow;
    /**
     *
     */
    private int lastRow;
    /**
     * 下拉单元格列号，从0开始
     */
    private int firstCol;
    /**
     *
     */
    private int lastCol;
    /**
     * 下拉单元格列（结束）
     */
    private int mergeNum;
    /**
     * 层级
     */
    private int level;
    /**
     * 数据集
     */
    private Map data;

    /**
     * 指定实体属性名
     */
    private String fieldName;
}
