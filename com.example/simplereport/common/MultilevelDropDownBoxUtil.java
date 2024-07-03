package com.example.simplereport.common;

import cn.hutool.core.util.ObjectUtil;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.springframework.util.CollectionUtils;

import java.util.*;
import java.util.stream.Collectors;


/**
 * 级联工具类
 */
public class MultilevelDropDownBoxUtil {
    public static void prepareGenerateMultilevelDropDownBox(Workbook wb, TreeMap<String, TreeMap<String, TreeMap<String, List<String>>>> dropDownDataSource,
                                                            List<String[]> allFieldsAndValues,List<MultiLevelDropDownContainer> containers) {
        generateMultilevelDropDownBox(wb,dropDownDataSource);
        Sheet sheet = wb.getSheetAt(0);
        //获取级联下拉值集合
        Map<String,Object> fieldsAndValuesMulti = new HashMap<>();
        List<String> fields = containers.stream().map(MultiLevelDropDownContainer::getFieldName).collect(Collectors.toList());
        allFieldsAndValues.stream().forEach(pair-> {
            if(fields.contains(pair[0])) fieldsAndValuesMulti.put(pair[0],pair[1]);
        });
        //赋值
        writeDropDownData(sheet,fieldsAndValuesMulti,containers);
    }

    /**
     * 生成联动四级下拉
     * @param wb
     * @param dropDownDataSource
     */
    public static void generateMultilevelDropDownBox(Workbook wb, TreeMap<String, TreeMap<String, TreeMap<String, List<String>>>> dropDownDataSource) {
        Sheet sheet = wb.getSheetAt(0);
        if (ObjectUtil.isEmpty(sheet)) {
            return;
        }

        String hiddenSheetName = "hiddensheet";
        Sheet hiddenSheet = wb.createSheet(hiddenSheetName);

        // 插入一级下拉框名称数据
        int rowIndex = 0;
        Row row = getOrCreateRow(hiddenSheet, rowIndex++);

        int colIndex = 0;
        Set<String> firstSelectNames = dropDownDataSource.keySet();
        String formulaName = "total" + firstSelectNames.size() + "firstSelect";
        row.createCell(colIndex).setCellValue(formulaName);
        for (String firstSelectName : firstSelectNames) {
            row = getOrCreateRow(hiddenSheet, rowIndex++);
            row.createCell(colIndex).setCellValue(firstSelectName);
        }

        String formulaName_firstSelect = formulaName;

        int endRowIndex = rowIndex;
        int beginRowIndex = rowIndex + 1;
        String colName = convertToExcelColumn(colIndex);
        String formula = hiddenSheetName + "!$" + colName + "$" + beginRowIndex + ":$" + colName + "$" + endRowIndex;
        bindFormula(wb, formulaName, formula);

        // 插入二级下拉框名称数据
        for (String firstSelectName : firstSelectNames) {
            TreeMap<String, TreeMap<String, List<String>>> map_secondSelectName_thirdSelectNames = dropDownDataSource.get(firstSelectName);
            Set<String> secondSelectNames = map_secondSelectName_thirdSelectNames.keySet();
            row = getOrCreateRow(hiddenSheet, rowIndex++);
            formulaName = firstSelectName;
            row.createCell(colIndex).setCellValue(formulaName);
            beginRowIndex = rowIndex + 1;
            for (String secondSelectName : secondSelectNames) {
                row = getOrCreateRow(hiddenSheet, rowIndex++);
                row.createCell(colIndex).setCellValue(secondSelectName);
            }
            endRowIndex = rowIndex;
            colName = convertToExcelColumn(colIndex);
            formula = hiddenSheetName + "!$" + colName + "$" + beginRowIndex + ":$" + colName + "$" + endRowIndex;
            bindFormula(wb, formulaName, formula);
        }

        // 插入三级下拉框名称数据
        for (String firstSelectName : firstSelectNames) {
            TreeMap<String, TreeMap<String, List<String>>> map_secondSelectName_thirdSelectNames = dropDownDataSource.get(firstSelectName);
            Set<String> secondSelectNames = map_secondSelectName_thirdSelectNames.keySet();
            for (String secondSelectName : secondSelectNames) {
                TreeMap<String, List<String>> map_thirdSelectName_thirdSelectNames = map_secondSelectName_thirdSelectNames.get(secondSelectName);
                Set<String> thirdSelectNames = map_thirdSelectName_thirdSelectNames.keySet();

                row = getOrCreateRow(hiddenSheet, rowIndex++);
                formulaName = secondSelectName;
                row.createCell(colIndex).setCellValue(formulaName);
                beginRowIndex = rowIndex + 1;
                for (String thirdSelectName : thirdSelectNames) {
                    row = getOrCreateRow(hiddenSheet, rowIndex++);
                    row.createCell(colIndex).setCellValue(thirdSelectName);
                }
                endRowIndex = rowIndex;
                colName = convertToExcelColumn(colIndex);
                formula = hiddenSheetName + "!$" + colName + "$" + beginRowIndex + ":$" + colName + "$" + endRowIndex;
                bindFormula(wb, formulaName, formula);
            }
        }

        // 插入四级下拉框名称数据
        for (String firstSelectName : firstSelectNames) {
            TreeMap<String, TreeMap<String, List<String>>> map_secondSelectName_fourSelectNames = dropDownDataSource.get(firstSelectName);
            Set<String> secondSelectNames = map_secondSelectName_fourSelectNames.keySet();
            for (String secondSelectName : secondSelectNames) {
                TreeMap<String, List<String>> map_thirdSelectName_fourSelectNames = map_secondSelectName_fourSelectNames.get(secondSelectName);
                Set<String> thirdSelectNames = map_thirdSelectName_fourSelectNames.keySet();
                for (String thirdSelectName : thirdSelectNames) {
                    List<String> fourSelectNames = map_thirdSelectName_fourSelectNames.get(thirdSelectName);
                    row = getOrCreateRow(hiddenSheet, rowIndex++);
                    formulaName = thirdSelectName;
                    row.createCell(colIndex).setCellValue(formulaName);
                    beginRowIndex = rowIndex + 1;
                    for (String fourSelectName : fourSelectNames) {
                        row = getOrCreateRow(hiddenSheet, rowIndex++);
                        row.createCell(colIndex).setCellValue(fourSelectName);
                    }
                    endRowIndex = rowIndex;
                    colName = convertToExcelColumn(colIndex);
                    formula = hiddenSheetName + "!$" + colName + "$" + beginRowIndex + ":$" + colName + "$" + endRowIndex;
                    bindFormula(wb, formulaName, formula);
                }
            }
        }

        DataValidationHelper helper = sheet.getDataValidationHelper();

        // 设置第1个下拉框
        generateSelectors(sheet, helper, formulaName_firstSelect, 9, 9, 2, 2);

        // 设置第2个下拉框
        colName = convertToExcelColumn(2);
        String indirect2 = "INDIRECT($" + colName + "1)";
        generateSelectors(sheet, helper, indirect2, 9, 9, 4, 4);

        // 设置第3个下拉框
        colName = convertToExcelColumn(4);
        String indirect3 = "INDIRECT($" + colName + "1)";
        generateSelectors(sheet, helper, indirect3, 9, 9, 6, 6);

        // 设置第4个下拉框
        colName = convertToExcelColumn(6);
        String indirect4 = "INDIRECT($" + colName + "1)";
        generateSelectors(sheet, helper, indirect4, 9, 9, 8, 8);


        //隐藏
        int sheetIndex = wb.getSheetIndex(hiddenSheetName);
        wb.setSheetHidden(sheetIndex, true);
    }

    /**
     * 给下拉框赋值
     * @param sheet
     * @param helper
     * @param indirect
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    private static void generateSelectors(Sheet sheet, DataValidationHelper helper, String indirect, int firstRow, int lastRow, int firstCol, int lastCol) {
        DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint(indirect);
        CellRangeAddressList region_firstSelect = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation validation_firstSelect = helper.createValidation(dvConstraint, region_firstSelect);
        if (validation_firstSelect instanceof XSSFDataValidation) {
            validation_firstSelect.setSuppressDropDownArrow(true);
            validation_firstSelect.setShowErrorBox(true);
        } else {
            validation_firstSelect.setSuppressDropDownArrow(false);
        }
        sheet.addValidationData(validation_firstSelect);
    }

    /**
     * @param colIndex 0 表示第 1 列
     * @return
     */
    private static String convertToExcelColumn(int colIndex) {
        StringBuilder s = new StringBuilder();
        while (colIndex >= 26) {
            s.insert(0, (char) ('A' + colIndex % 26));
            colIndex = colIndex / 26 - 1;
        }
        s.insert(0, (char) ('A' + colIndex));
        return s.toString();
    }

    private static void bindFormula(Workbook workbook, String formulaName, String formula) {
        Name name = workbook.createName();
        name.setNameName(formulaName);
        name.setRefersToFormula(formula);
        //  System.out.println("bind \"" + formulaName + "\" for formula \"" + formula + "\"");
    }

    private static Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            return row;
        }
        return sheet.createRow(rowIndex);
    }

    /**
     * 给下拉框设置选中值
     * @param sheet
     * @param fieldAndValue
     * @param multiLevelDropDownContainers
     */
    public static void writeDropDownData(Sheet sheet,Map<String,Object> fieldAndValue,List<MultiLevelDropDownContainer> multiLevelDropDownContainers){
        if(CollectionUtils.isEmpty(fieldAndValue) || CollectionUtils.isEmpty(multiLevelDropDownContainers)){
            return;
        }
        multiLevelDropDownContainers.stream().forEach(dto->{
            Object value = fieldAndValue.get(dto.getFieldName());
            Cell cell = sheet.getRow(dto.getFirstRow()).getCell(dto.getFirstCol());
            cell.setCellValue(String.valueOf(value));
            cell.setCellStyle(cell.getCellStyle());
        });
    }
}