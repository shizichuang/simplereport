package com.example.simplereport.common;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.example.simplereport.entity.SurveySimpleReport;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.springframework.util.CollectionUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
public class EasyPoiUtil {
    private static String templateName = "/templates/调查简表模板.xls";

    /**
     * 导出
     * @param simpleReport 简单报表对象
     * @param selectors 定向下拉框集合
     * @param blockAreaContainers 块状对象容器集合--主要为了不规则块状对象位置、值得预定义
     * @return
     */
    public static Workbook prepareExportExcel(SurveySimpleReport simpleReport, List<ExcelSelectorContainer> selectors, List<BlockAreaContainer> blockAreaContainers) {
        Workbook wb = null;
        try {
            //非规则嵌套得块状区域对象集合
            Map<String,List> notRuleEntityMap = CommonUtil.getEntityList(simpleReport,false);
            //规则嵌套得块状区域对象集合
            Map<String,List> ruleEntityMap = CommonUtil.getEntityList(simpleReport,true);
            //级联数据集
            TreeMap<String,TreeMap<String,TreeMap<String, List<String>>>> treeMap = CommonUtil.dropDownDataSource();
            //级联下拉框容器集
            List<MultiLevelDropDownContainer> MContainers = CommonUtil.getMultiLevelDropDownContainers();
            //导出
            wb = exportExcel(simpleReport,selectors, blockAreaContainers,notRuleEntityMap,ruleEntityMap,treeMap,MContainers);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            log.error("反射获取值异常");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            log.error("转码异常");
        }
        return wb;
    }
    public static Workbook exportExcel(SurveySimpleReport simpleReport, List<ExcelSelectorContainer> selectors, List<BlockAreaContainer> blockAreaContainers, Map<String,List> noRuleEntityMap, Map<String,List> ruleEntityMap,
                                       TreeMap treeMap,List<MultiLevelDropDownContainer> MContainers) throws IllegalAccessException, UnsupportedEncodingException {
        //获取模板URL
        String templatePath = EasyPoiUtil.class.getClass().getResource(templateName).getPath();
        // 对路径进行解码，这里假设资源文件名使用的是UTF-8编码
        templatePath = URLDecoder.decode(templatePath, StandardCharsets.UTF_8.name());

        // 创建TemplateExportParams实例，传入模板文件名（这里可以省略，因为实际内容是从流中读取）
        TemplateExportParams templateExportParams = new TemplateExportParams(templatePath, 0);
        //开启横向遍历 开启横向遍历 开启横向遍历
        templateExportParams.setColForEach(true);

        //下拉框默认值单独定义，存放在这个集合中，通过fieldName从simTable对象集合中
        Map<String,Object> fieldsAndValuesSelector = new HashMap<>();
        //获取所有下拉框得字段名称
        List<String> fields = selectors.stream().map(ExcelSelectorContainer::getFieldName).collect(Collectors.toList());

        //不包含下拉框得字段值集合，直接写入excel，让模板取值
        Map<String, Object> fieldsAndValuesNoSector = new HashMap<>();
        List<String[]> allFieldsAndValues  = CommonUtil.getAllFieldsAndValues(simpleReport);
        allFieldsAndValues.stream().forEach(pair-> {
            if(fields.contains(pair[0])) fieldsAndValuesSelector.put(pair[0],pair[1]);
            else fieldsAndValuesNoSector.put(pair[0], pair[1]);
        });
        //非规则嵌套块状区域对象,集合放入页面参数中通过模板方法遍历赋值，注意放开前注释掉特殊块状对象遍历
        //applicationInfoMapNoSector.putAll(noRuleEntityMap);

        //生成对象
        Workbook workbook = null;
        try {
            workbook = ExcelExportUtil.exportExcel(templateExportParams, fieldsAndValuesNoSector);
            //特殊块状对象处理-评估信息、申请人关联企业
            prepareWriteSpecialBlockObject(workbook, blockAreaContainers,noRuleEntityMap);
            //规则块集合纵向遍历加设置动态下拉框
            prepareVerticalAssignment(workbook, blockAreaContainers,ruleEntityMap,selectors);
            //预生成单个的不连续的下拉框并赋值
            prepareGenerateSingleDropDownBoxAndSetValue(workbook,selectors,fieldsAndValuesSelector);
            //四级联动设置
            MultilevelDropDownBoxUtil.prepareGenerateMultilevelDropDownBox(workbook,treeMap,allFieldsAndValues,MContainers);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return workbook;
    }

    /**
     * 纵向遍历
     * @param workbook
     * @param blockAreaContainers 纵向遍历块容器，部分值要提前定义好
     * @param listMap<classNme,List<class>/> 这里获取的待遍历对象的集合
     * @return
     */
    public static Workbook prepareVerticalAssignment(Workbook workbook, List<BlockAreaContainer> blockAreaContainers, Map<String,List> listMap, List<ExcelSelectorContainer> selectors)  {
        if(CollectionUtils.isEmpty(blockAreaContainers) || CollectionUtils.isEmpty(listMap)){
            return workbook;
        }
        Sheet sheet = workbook.getSheetAt(0);
        blockAreaContainers.stream().forEach(blockAreaContainer -> {
            String objectName = blockAreaContainer.getObjectName();
            List<ExcelSelectorContainer> selectorList = selectors.stream().filter(selector->selector.getFieldName().contains(objectName)).collect(Collectors.toList());
            blockAreaContainer.setSelectorDtoList(selectorList);
            blockAreaContainer.setDataList(listMap.get(blockAreaContainer.getObjectName()));
            try {
                verticalAssignment(sheet, blockAreaContainer);
            } catch (IllegalAccessException e) {
                e.printStackTrace();

            }
        });
        return workbook;
    }

    /**
     * 纵向遍历写入值
     * @param sheet
     * @param container
     * @throws IllegalAccessException
     */
    public static void verticalAssignment(Sheet sheet, BlockAreaContainer container) throws IllegalAccessException {
        int baseRowIndex = container.getBaseRowIndex();
        int baseColIndex = container.getBaseColIndex();
        int mergeColIndex = container.getMergeColIndex();
        List<ExcelSelectorContainer> selectorDtoList = container.getSelectorDtoList();

        List blockObjectList = container.getDataList();
        if( CollectionUtils.isEmpty(blockObjectList)){
            return;
        }

        ArrayList arrayList = (ArrayList) blockObjectList.get(0);

        for (int i = 0; i < arrayList.size(); i++) {
            Row row = sheet.getRow(baseRowIndex + i);
            // 获取每一行的列值集合
            Field[] fields = arrayList.get(i).getClass().getDeclaredFields();
            for (int j = 0; j < fields.length; j++) {
                Field field = fields[j];
                field.setAccessible(true); // 允许访问私有字段

                String fieldName = field.getName();
                List<ExcelSelectorContainer> selectors = selectorDtoList.stream().filter(ite->ite.getFieldName().contains(fieldName)).collect(Collectors.toList());
                if (!CollectionUtils.isEmpty(selectors)) {
                    ExcelSelectorContainer selector = selectors.get(0);
                    selector = CommonUtil.setSelector(selector.getFieldName(),baseRowIndex + i,baseColIndex + j * mergeColIndex,selector.getDatas(),true);
                    selectors = new ArrayList<>();
                    selectors.add(selector);
                    generateDropDownBox(sheet,selectors);
                }

                Object fieldValue = field.get(arrayList.get(i));
                Cell cell = row.getCell(baseColIndex + j * mergeColIndex);
                cell.setCellValue(String.valueOf(fieldValue==null?"":fieldValue));
                cell.setCellStyle(cell.getCellStyle());
            }
        }
    }

    /**
     * 不规则特殊块状区域预写入
     * @param workbook
     * @param blockAreaContainers
     * @param listMap
     * @return
     */
    public static Workbook prepareWriteSpecialBlockObject(Workbook workbook, List<BlockAreaContainer> blockAreaContainers, Map<String,List> listMap)  {
        if(CollectionUtils.isEmpty(blockAreaContainers) || CollectionUtils.isEmpty(listMap)){
            return workbook;
        }
        Sheet sheet = workbook.getSheetAt(0);
        blockAreaContainers.stream().forEach(blockAreaContainer -> {
            blockAreaContainer.setDataList(listMap.get(blockAreaContainer.getObjectName()));
            try {
                writeSpecialBlockObject(sheet, blockAreaContainer);
            } catch (IllegalAccessException e) {
                e.printStackTrace();

            }
        });
        return workbook;
    }

    /**
     * 不规则特殊块状区域写入
     * @param sheet
     * @param container
     * @throws IllegalAccessException
     */
    public static void writeSpecialBlockObject(Sheet sheet, BlockAreaContainer container) throws IllegalAccessException {
        int baseRowIndex = container.getBaseRowIndex();
        int baseColIndex = container.getBaseColIndex();
        List blockObjectList = container.getDataList();
        if(CollectionUtils.isEmpty(blockObjectList)) return;

        ArrayList arrayList = (ArrayList) blockObjectList.get(0);
        for (int i = 0; i < arrayList.size(); i++) {
            Row row = sheet.getRow(baseRowIndex + i);

            Field[] fields = arrayList.get(i).getClass().getDeclaredFields();

            //
            Field field = fields[0];
            field.setAccessible(true); // 允许访问私有字段
            Cell cell1 = row.getCell(baseColIndex);
            cell1.setCellValue(String.valueOf(field.get(arrayList.get(i))));

            Field field1 = fields[1];
            field1.setAccessible(true); // 允许访问私有字段
            Cell cell2 = row.getCell(baseColIndex+3);
            cell2.setCellValue(String.valueOf(field1.get(arrayList.get(i))));

            Field field2 = fields[2];
            field2.setAccessible(true); // 允许访问私有字段
            Cell cell3 = row.getCell(baseColIndex+5);
            cell3.setCellValue(String.valueOf(field2.get(arrayList.get(i))));
        }
    }

    /**
     * 预生成单个的下拉框并赋值
     * @param workbook
     * @param selectors
     * @param selectorFieldsAndValues
     */
    public static void prepareGenerateSingleDropDownBoxAndSetValue(Workbook workbook,List<ExcelSelectorContainer> selectors, Map<String,Object> selectorFieldsAndValues) {
        Sheet sheet = workbook.getSheetAt(0);
        selectors = selectors.stream().filter(s->s.isForForeach()==false).collect(Collectors.toList());
        generateDropDownBox(sheet,selectors);
        writeDropDownData(sheet,selectorFieldsAndValues,selectors);
        generateDropDownBox(sheet,selectors);
    }
    /**
     * 生成下拉框并预设值
     * @param sheet
     * @param selectorDtos
     * @return
     */
    public static void generateDropDownBox(Sheet sheet,List<ExcelSelectorContainer> selectorDtos){
        if(CollectionUtils.isEmpty(selectorDtos)){
            return;
        }
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList cellRangeAddressList;
        DataValidationConstraint dvConstraint;
        DataValidation dataValidation;
        for (ExcelSelectorContainer dto : selectorDtos) {
            cellRangeAddressList = new CellRangeAddressList(dto.getFirstRow(),dto.getLastRow(),dto.getFirstCol(),dto.getLastCol());
            //生成下拉框内容
            dvConstraint = DVConstraint.createExplicitListConstraint(dto.getDatas());
            dataValidation = helper.createValidation(dvConstraint,cellRangeAddressList);
            // Excel兼容性问题
            if (dataValidation instanceof XSSFDataValidation) {
                dataValidation.setSuppressDropDownArrow(true);
                dataValidation.setShowErrorBox(true);
            } else {
                dataValidation.setSuppressDropDownArrow(false);
            }
            //对sheet页生效
            sheet.addValidationData(dataValidation);
        }
    }

    /**
     * 给下拉框设置选中值
     * @param sheet
     * @param selectorFieldsAndValues
     * @param selectorDtos
     */
    public static void writeDropDownData(Sheet sheet,Map<String,Object> selectorFieldsAndValues,List<ExcelSelectorContainer> selectorDtos){
        if(CollectionUtils.isEmpty(selectorFieldsAndValues) || CollectionUtils.isEmpty(selectorDtos)){
            return;
        }
        selectorDtos.stream().forEach(dto->{
            Object value = selectorFieldsAndValues.get(dto.getFieldName());
            Cell cell = sheet.getRow(dto.getFirstRow()).getCell(dto.getFirstCol());
            cell.setCellValue(String.valueOf(value));
            cell.setCellStyle(cell.getCellStyle());
        });
    }

    /**
     * 下载
     *
     * @param fileName 文件名称
     * @param response
     * @param workbook excel数据
     */
    public static void downLoadExcel(String fileName, HttpServletResponse response, Workbook workbook) throws Exception {
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            throw new Exception(e.getMessage());
        }
    }

    public static void simpleReportExport(HttpServletResponse response,List<ValidErrorCell> errorCells) throws Exception {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams(), ValidErrorCell.class, errorCells);
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + URLEncoder.encode("错误信息", "UTF-8"));
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            throw new Exception(e.getMessage());
        }
    }
}
