package com.example.simplereport.common;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.hutool.core.collection.CollectionUtil;
import com.alibaba.fastjson.JSONObject;
import com.example.simplereport.entity.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;

public class PoiUtil {
    /**
     * 导入excel
     */
    public static void inportExcel(File file, HttpServletResponse response){
        Workbook workbook=null;
        try {
            workbook = WorkbookFactory.create(file);
            //获取第一个sheet页,根据模板单页处理
            Sheet sheet = workbook.getSheetAt(0);
            //获取所有合并单元格
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

            SurveySimpleReport surveySimpleReport = new SurveySimpleReport();//调查简单报告-总
            ApplicantBaseInfo applicantBaseInfo = new ApplicantBaseInfo();//申请人进件基础信息
            ApplicantBusinessInfo applicantBusinessInfo = new ApplicantBusinessInfo();//申请人业务信息
            CollateralInfo collateralInfo = new CollateralInfo();//担保信息（抵押）
            ApplicantReportBaseInfo applicantReportBaseInfo = new ApplicantReportBaseInfo();//申请人调查报告-基本信息
            ApplicantReportOperationalInformation applicantReportOperationalInformation = new ApplicantReportOperationalInformation();//申请人调查报告-申请人及关联主体经营信息分析
            ApplicantReportCreditAnalysisConclusion applicantReportCreditAnalysisConclusion = new ApplicantReportCreditAnalysisConclusion();//申请人调查报告-申请人及关联主体征信分析结论
            MortgageAssetOtherInfo mortgageAssetOtherInfo = new MortgageAssetOtherInfo();//抵押物其他信息
            BusinessCriticalIndicator businessCriticalIndicator = new BusinessCriticalIndicator();//业务重要标识
            CreditInquiryParent creditInquiryParent = new CreditInquiryParent();//征信解析
            CreditExposureCalculationNew creditExposureCalculationNew = new CreditExposureCalculationNew();//授信敞口测算（新）
            ElectronicSurveyReport electronicSurveyReport = new ElectronicSurveyReport();//电子版调查报告
            String signingOpinion = "";//签署意见

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                String value = "";//单元格值
                int firstRow = 0;//起始行
                int lastRow = 0;//结束行
                Row row = sheet.getRow( i);
                Cell cell = row.getCell(0);//首列单元格
                //空值不处理
                if (StringUtils.isEmpty(getStringValue(cell))){
                    continue;
                }
                //判断当前单元格是否是‘合并单元格’
                CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, cell);
                if (Objects.nonNull(cellRangeAddress)) {
                    //当前首列不是空值，并且是合并单元格里边的
                    value = getCellRangeAddressValue(cellRangeAddress, sheet);
                    firstRow = cellRangeAddress.getFirstRow();
                    lastRow = cellRangeAddress.getLastRow();
                } else {
                    //空值；或者是单独一行的模块，例如签署意见
                    value = getStringValue(cell);
                    firstRow = i;
                    lastRow = i;
                }
                value = value.trim();
                /**
                 * 根据模板解析，可定义为以下几个模块
                 * 处理方案：解析excel所有的合并单元格，并根据首列匹配上述模块，分模块解析落地
                 *          模块名称，模块起始行，模块结束行。
                 * 说明：部分模块非简单key-value格式，且存在条数动态变化情况，如有调整，基于模块个性化处理
                 */

                if ("申请人进件基础信息".equals(value)){
                    analysisExcelCommon(firstRow, lastRow ,sheet, applicantBaseInfo);
                }else if ("申请人业务信息".equals(value)){
                    applicantBusinessInfo = analysisExcelBusinessInfo(firstRow, lastRow ,sheet);
                }else if ("担保信息（抵押）".equals(value)){
                    analysisExcelCommon(firstRow, lastRow ,sheet, collateralInfo);
                }else if ("申请人调查报告-基本信息".equals(value)){
                    analysisExcelCommon(firstRow, lastRow ,sheet, applicantReportBaseInfo);
                }else if ("申请人调查报告-申请人及关联主体经营信息分析".equals(value)){
                    analysisExcelCommon(firstRow, lastRow ,sheet, applicantReportOperationalInformation);
                }else if ("申请人调查报告-申请人及关联主体征信分析结论".equals(value)){
                    analysisExcelCommon(firstRow, lastRow ,sheet, applicantReportCreditAnalysisConclusion);
                }else if ("抵押物其他信息".equals(value)){
                    mortgageAssetOtherInfo = analysisExcelMortgageAssetOtherInfo(firstRow, lastRow ,sheet);
                }else if ("业务重要标识".equals(value)){
                    analysisExcelCommon(firstRow, lastRow ,sheet, businessCriticalIndicator);
                }else if ("征信解析".equals(value)){
                    creditInquiryParent = analysisExcelCreditInquiryParent(firstRow, lastRow ,sheet);
                }else if ("授信敞口测算（新）".equals(value)){
                    analysisExcelCommon(firstRow, lastRow ,sheet, creditExposureCalculationNew);
                }else if ("电子版调查报告".equals(value)){
                    electronicSurveyReport = analysisExcelElectronicSurveyReport(firstRow, lastRow ,sheet);
                }else if ("签署意见".equals(value)){
                    signingOpinion = analysisExcelOpinion(firstRow, lastRow ,sheet);
                }

            }
            //打印解析内容
            surveySimpleReport.setApplicantBaseInfo(applicantBaseInfo);
            surveySimpleReport.setApplicantBusinessInfo(applicantBusinessInfo);
            surveySimpleReport.setCollateralInfo(collateralInfo);
            surveySimpleReport.setApplicantReportBaseInfo(applicantReportBaseInfo);
            surveySimpleReport.setApplicantReportOperationalInformation(applicantReportOperationalInformation);
            surveySimpleReport.setAnalysisConclusion(applicantReportCreditAnalysisConclusion);
            surveySimpleReport.setAssetOtherInfo(mortgageAssetOtherInfo);
            surveySimpleReport.setBusinessCriticalIndicator(businessCriticalIndicator);
            surveySimpleReport.setCreditInquiryP(creditInquiryParent);
            surveySimpleReport.setCreditExposureCalculationNew(creditExposureCalculationNew);
            surveySimpleReport.setElectronicSurveyReport(electronicSurveyReport);
            surveySimpleReport.setSigningOpinion(signingOpinion);

            //错误信息导出
            List<ValidErrorCell> errorCellList = new ArrayList<>();
            errorCellList = CommonUtil.validTableColumn(errorCellList,electronicSurveyReport);
            if(!CollectionUtils.isEmpty(errorCellList)){
                EasyPoiUtil.simpleReportExport(response,errorCellList);
            }
            //实体类入库
            String json = JSONObject.toJSON(surveySimpleReport).toString();
            System.out.println(json);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e1) {
            e1.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (Exception e) {
                workbook=null;
                e.printStackTrace();
            }
        }
    }

    /**
     * 解析key-value(合并单元格也能正确处理)标准格式的模块，例如：“申请人进件基础信息”
     * @param firstRow 起始行
     * @param lastRow 结束行
     * @param sheet excel页面
     * @param object 保存解析内容的对象
     * @return ApplicantBaseInfo
     */
    private static void analysisExcelCommon(int firstRow, int lastRow, Sheet sheet, Object object){
        Map<String, Object> map = new HashMap<>();
        for (int i=firstRow; i<=lastRow; i++){
            Row row = sheet.getRow(i);
            //注意剔除首行， j=1
            for (int j=1; j<row.getLastCellNum(); j++){
                //如果当前单元格值等于类的注解（遍历判断），当前申请人信息默认读取下一个单元格值（横向）
                String attributeValue = checkExcelAttribute(object, row.getCell(j));
                if (StringUtils.isNotEmpty(attributeValue)){
                    map.put(attributeValue, getStringValue(row.getCell(j+1)));
                }
            }
        }
        //将map值赋值到对象中去
        if (!map.isEmpty()){
            setMapParam(object, map);
        }
    }
//    /**
//     * 解析“申请人进件基础信息”
//     * @param firstRow 起始行
//     * @param lastRow 结束行
//     * @param sheet
//     * @return ApplicantBaseInfo
//     */
//    private static ApplicantBaseInfo analysisExcelBaseInfo(int firstRow, int lastRow, Sheet sheet){
//        ApplicantBaseInfo applicantBaseInfo = new ApplicantBaseInfo();
//        Map<String, Object> map = new HashMap<>();
//        for (int i=firstRow; i<lastRow; i++){
//            Row row = sheet.getRow(i);
//            //注意剔除首行， j=1
//            for (int j=1; j<row.getLastCellNum(); j++){
//                //如果当前单元格值等于类的注解（遍历判断），当前申请人信息默认读取下一个单元格值（横向）
//                String attributeValue = checkExcelAttribute(applicantBaseInfo, row.getCell(j));
//                if (StringUtils.isNotEmpty(attributeValue)){
//                    map.put(attributeValue, getStringValue(row.getCell(j+1)));
//                }
//            }
//        }
//        //将map值赋值到申请人对象中去
//        if (!map.isEmpty()){
//            setMapParam(applicantBaseInfo, map);
//        }
//        return applicantBaseInfo;
//    }
    /**
     * 解析“申请人业务信息”。 特殊处理
     * @param firstRow 起始行
     * @param lastRow 结束行
     * @param sheet
     * @return ApplicantBusinessInfo
     */
    private static ApplicantBusinessInfo analysisExcelBusinessInfo(int firstRow, int lastRow, Sheet sheet){
        ApplicantBusinessInfo applicantBusinessInfo = new ApplicantBusinessInfo();
        Map<String, Object> map = new HashMap<>();
        for (int i=firstRow; i<=lastRow; i++){
            Row row = sheet.getRow(i);
            //注意剔除首行， j=1
            for (int j=1; j<row.getLastCellNum(); j++){
                //如果当前单元格值等于类的注解（遍历判断），当前申请人信息默认读取下一个单元格值（横向）
                String attributeValue = checkExcelAttribute(applicantBusinessInfo, row.getCell(j));
                if (StringUtils.isNotEmpty(attributeValue)){
                    map.put(attributeValue, getStringValue(row.getCell(j+1)));
                }
                //申请人业务信息（小企业特色业务（可多选），行业投向）单独处理。 注意ApplicantBusinessInfo不要对这俩字段加注解
                if ("小企业特色业务（可多选）".equals(getStringValue(row.getCell(j)))){
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+1)))){
                        applicantBusinessInfo.setSmallBusinessFeatures1(getStringValue(row.getCell(j+1)));
                    }
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+2)))){
                        applicantBusinessInfo.setSmallBusinessFeatures1(getStringValue(row.getCell(j+2)));
                    }
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+3)))){
                        applicantBusinessInfo.setSmallBusinessFeatures1(getStringValue(row.getCell(j+3)));
                    }
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+4)))){
                        applicantBusinessInfo.setSmallBusinessFeatures1(getStringValue(row.getCell(j+4)));
                    }
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+5)))){
                        applicantBusinessInfo.setSmallBusinessFeatures1(getStringValue(row.getCell(j+5)));
                    }
                }
                if ("行业投向".equals(getStringValue(row.getCell(j)))){
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+1)))){
                        applicantBusinessInfo.setIndustryOrientation1(getStringValue(row.getCell(j+1)));
                    }
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+3)))){
                        applicantBusinessInfo.setIndustryOrientation2(getStringValue(row.getCell(j+3)));
                    }
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+5)))){
                        applicantBusinessInfo.setIndustryOrientation3(getStringValue(row.getCell(j+5)));
                    }
                    if (StringUtils.isNotEmpty(getStringValue(row.getCell(j+7)))){
                        applicantBusinessInfo.setIndustryOrientation4(getStringValue(row.getCell(j+7)));
                    }
                }
            }
        }
        //将map值赋值到申请人对象中去
        if (!map.isEmpty()){
            setMapParam(applicantBusinessInfo, map);
        }
        return applicantBusinessInfo;
    }
    /**
     * 解析“抵押物其他信息”。 特殊处理. key-value的还是通过注解读取。其他的个性化读取
     * @param firstRow 起始行
     * @param lastRow 结束行
     * @param sheet
     * @return MortgageAssetOtherInfo
     */
    private static MortgageAssetOtherInfo analysisExcelMortgageAssetOtherInfo(int firstRow, int lastRow, Sheet sheet){
        MortgageAssetOtherInfo mortgageAssetOtherInfo = new MortgageAssetOtherInfo();
        Map<String, Object> map = new HashMap<>();
        for (int i=firstRow; i<=lastRow; i++){
            Row row = sheet.getRow(i);
            //注意剔除首行， j=1
            for (int j=1; j<row.getLastCellNum(); j++){
                //如果当前单元格值等于类的注解（遍历判断），当前申请人信息默认读取下一个单元格值（横向）
                //
                String attributeValue = checkExcelAttribute(mortgageAssetOtherInfo, row.getCell(j));
                if (StringUtils.isNotEmpty(attributeValue)){
                    map.put(attributeValue, getStringValue(row.getCell(j+1)));
                }
                //评估信息（最多维护3条）单独处理。 注意MortgageAssetOtherInfo不要对这个字段加注解
                if ("评估信息（最多维护3条）".equals(getStringValue(row.getCell(j)))){
                    //开始判断“评估信息（最多维护3条）”占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<AppraisalInformation> appraisalInformationList = new ArrayList<>();
                    int assetfirstRow = 0;
                    int assetlastRow = 0;
                    int assetlastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        assetfirstRow = cellRangeAddress.getFirstRow();
                        assetlastRow = cellRangeAddress.getLastRow();
                        assetlastColumn = cellRangeAddress.getLastColumn();
                        for (int s= assetfirstRow+1; s<assetlastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(assetlastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            AppraisalInformation appraisalInformation = new AppraisalInformation();//评估信息（最多维护3条）
                            //评估信息来源
                            appraisalInformation.setAppraisalSource(getStringValue(sheet.getRow(s).getCell(assetlastColumn+1)));
                            //评估单价(元)  加3列
                            String appraisalUnitPrice = getStringValue(sheet.getRow(s).getCell(assetlastColumn+4));
                            if (StringUtils.isNotEmpty(appraisalUnitPrice)){
                                appraisalInformation.setAppraisalUnitPrice(new BigDecimal(appraisalUnitPrice));
                            }
                            //评估总价(元) 加5列
                            String appraisalTotalPrice = getStringValue(sheet.getRow(s).getCell(assetlastColumn+6));
                            if (StringUtils.isNotEmpty(appraisalTotalPrice)){
                                appraisalInformation.setAppraisalTotalPrice(new BigDecimal(appraisalTotalPrice));
                            }
                            //加入集合
                            appraisalInformationList.add(appraisalInformation);
                        }
                    }
                    //设置评估信息（最多维护3条）
                    mortgageAssetOtherInfo.setAppraisalInformations(appraisalInformationList);
                }
            }
        }
        //将map值赋值到申请人对象中去
        if (!map.isEmpty()){
            setMapParam(mortgageAssetOtherInfo, map);
        }
        return mortgageAssetOtherInfo;
    }
    /**
     * 解析“征信解析”。 特殊处理. key-value的还是通过注解读取。其他的个性化读取
     * @param firstRow 起始行
     * @param lastRow 结束行
     * @param sheet
     * @return CreditInquiryParent
     */
    private static CreditInquiryParent analysisExcelCreditInquiryParent(int firstRow, int lastRow, Sheet sheet){
        CreditInquiryParent creditInquiryParent = new CreditInquiryParent();
        Map<String, Object> map = new HashMap<>();
        for (int i=firstRow; i<=lastRow; i++){
            Row row = sheet.getRow(i);
            //注意剔除首行， j=1
            for (int j=1; j<row.getLastCellNum(); j++){
                //如果当前单元格值等于类的注解（遍历判断），当前申请人信息默认读取下一个单元格值（横向）
                //
                String attributeValue = checkExcelAttribute(creditInquiryParent, row.getCell(j));
                if (StringUtils.isNotEmpty(attributeValue)){
                    map.put(attributeValue, getStringValue(row.getCell(j+1)));
                }
                //征信解析（包括但不限于经营背景、关联企业、其他个人等，最多维护6条）单独处理。 注意CreditInquiryParent不要对这个字段加注解
                if ("征信解析（包括但不限于经营背景、关联企业、其他个人等，最多维护6条）".equals(getStringValue(row.getCell(j)))){
                    //开始判断“征信解析（包括但不限于经营背景、关联企业、其他个人等，最多维护6条）”占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<CreditInquiry> creditInquiryList = new ArrayList<>();
                    int mergedfirstRow = 0;
                    int mergedlastRow = 0;
                    int mergedlastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        mergedfirstRow = cellRangeAddress.getFirstRow();
                        mergedlastRow = cellRangeAddress.getLastRow();
                        mergedlastColumn = cellRangeAddress.getLastColumn();
                        for (int s= mergedfirstRow+1; s<mergedlastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(mergedlastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            CreditInquiry creditInquiry = new CreditInquiry();//征信解析（包括但不限于经营背景、关联企业、其他个人等，最多维护6条）
                            //查询对象类型
                            creditInquiry.setInquiryObjectType(getStringValue(sheet.getRow(s).getCell(mergedlastColumn+1)));
                            //评客户名称  加2列
                            creditInquiry.setCustomerName(getStringValue(sheet.getRow(s).getCell(mergedlastColumn+3)));
                            //证件号码  加4列
                            creditInquiry.setIdNumber(getStringValue(sheet.getRow(s).getCell(mergedlastColumn+5)));
                            //查询证件号码  加6列
                            creditInquiry.setQueryIdNumber(getStringValue(sheet.getRow(s).getCell(mergedlastColumn+7)));
                            //加入集合
                            creditInquiryList.add(creditInquiry);
                        }
                    }
                    //设置征信解析（包括但不限于经营背景、关联企业、其他个人等，最多维护6条）
                    creditInquiryParent.setCreditInquiries(creditInquiryList);
                }
            }
        }
        //将map值赋值到申请人对象中去
        if (!map.isEmpty()){
            setMapParam(creditInquiryParent, map);
        }
        return creditInquiryParent;
    }
    /**
     * 解析“电子版调查报告”。 特殊处理. key-value的还是通过注解读取。其他的个性化读取
     * @param firstRow 起始行
     * @param lastRow 结束行
     * @param sheet
     * @return ElectronicSurveyReport
     */
    private static ElectronicSurveyReport analysisExcelElectronicSurveyReport(int firstRow, int lastRow, Sheet sheet){
        ElectronicSurveyReport electronicSurveyReport = new ElectronicSurveyReport();
        Map<String, Object> map = new HashMap<>();
        for (int i=firstRow; i<=lastRow; i++){
            Row row = sheet.getRow(i);
            //注意剔除首行， j=1
            for (int j=1; j<row.getLastCellNum(); j++){
                //如果当前单元格值等于类的注解（遍历判断），当前申请人信息默认读取下一个单元格值（横向）
                //
                String attributeValue = checkExcelAttribute(electronicSurveyReport, row.getCell(j));
                if (StringUtils.isNotEmpty(attributeValue)){
                    map.put(attributeValue, getStringValue(row.getCell(j+1)));
                }
                //借款主体/实际控制人近三年从业信息（最多维护3条）单独处理。 注意ElectronicSurveyReport不要对这个字段加注解
                if ("借款主体/实际控制人近三年从业信息（最多维护3条）".equals(getStringValue(row.getCell(j)))){
                    //开始判断“借款主体/实际控制人近三年从业信息（最多维护3条）”占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<EmploymentHistory> employmentHistoryList = new ArrayList<>();
                    int subfirstRow = 0;
                    int sublastRow = 0;
                    int sublastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        subfirstRow = cellRangeAddress.getFirstRow();
                        sublastRow = cellRangeAddress.getLastRow();
                        sublastColumn = cellRangeAddress.getLastColumn();
                        for (int s= subfirstRow+1; s<sublastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            EmploymentHistory employmentHistory = new EmploymentHistory();//借款主体/实际控制人近三年从业信息（最多维护3条）
                            //开始时间
                            String startDate = getStringValue(sheet.getRow(s).getCell(sublastColumn+1));
                            if (StringUtils.isNotEmpty(startDate)){
                                employmentHistory.setStartDate(new Date(startDate));
                            }
                            //结束时间  加2列
                            String endDate = getStringValue(sheet.getRow(s).getCell(sublastColumn+3));
                            if (StringUtils.isNotEmpty(endDate)){
                                employmentHistory.setEndDate(new Date(endDate));
                            }
                            //工作单位名称  加4列
                            employmentHistory.setEmployerName(getStringValue(sheet.getRow(s).getCell(sublastColumn+5)));
                            //岗位或职务 加6列
                            employmentHistory.setPositionOrDuty(getStringValue(sheet.getRow(s).getCell(sublastColumn+7)));
                            //加入集合
                            employmentHistoryList.add(employmentHistory);
                        }
                    }
                    //设置借款主体/实际控制人近三年从业信息（最多维护3条）
                    electronicSurveyReport.setEmploymentHistories(employmentHistoryList);
                }
                //经营背景股权结构（最多维护4条）单独处理。 注意ElectronicSurveyReport不要对这个字段加注解
                if ("经营背景股权结构（最多维护4条）".equals(getStringValue(row.getCell(j)))){
                    //开始判断“经营背景股权结构（最多维护4条）”占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<EquityStructure> equityStructureList = new ArrayList<>();
                    int subfirstRow = 0;
                    int sublastRow = 0;
                    int sublastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        subfirstRow = cellRangeAddress.getFirstRow();
                        sublastRow = cellRangeAddress.getLastRow();
                        sublastColumn = cellRangeAddress.getLastColumn();
                        for (int s= subfirstRow+1; s<sublastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            EquityStructure equityStructure = new EquityStructure();//经营背景股权结构（最多维护4条）
                            //股东名称
                            equityStructure.setShareholderName(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)));
                            //实际投资金额(元)  加2列
                            String actualInvestmentAmount = getStringValue(sheet.getRow(s).getCell(sublastColumn+3));
                            if (StringUtils.isNotEmpty(actualInvestmentAmount)){
                                equityStructure.setActualInvestmentAmount(new BigDecimal(actualInvestmentAmount));
                            }
                            //出资比例(%)  加4列
                            String contributionRatio = getStringValue(sheet.getRow(s).getCell(sublastColumn+5));
                            if (StringUtils.isNotEmpty(contributionRatio)){
                                equityStructure.setContributionRatio(Double.valueOf(contributionRatio));
                            }
                            //是否为法人 加6列
                            String isLegalPerson = getStringValue(sheet.getRow(s).getCell(sublastColumn+7));
                            if (StringUtils.isNotEmpty(isLegalPerson)){
                                equityStructure.setIsLegalPerson(Boolean.valueOf(isLegalPerson));
                            }
                            //加入集合
                            equityStructureList.add(equityStructure);
                        }
                    }
                    //设置经营背景股权结构（最多维护4条）
                    electronicSurveyReport.setEquityStructures(equityStructureList);
                }
                //申请人关联企业单独处理。 注意ElectronicSurveyReport不要对这个字段加注解
                if ("申请人关联企业".equals(getStringValue(row.getCell(j)))){
                    //开始判断“申请人关联企业”占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<ApplicantRelatedCompany> applicantRelatedCompanyList = new ArrayList<>();
                    int subfirstRow = 0;
                    int sublastRow = 0;
                    int sublastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        subfirstRow = cellRangeAddress.getFirstRow();
                        sublastRow = cellRangeAddress.getLastRow();
                        sublastColumn = cellRangeAddress.getLastColumn();
                        for (int s= subfirstRow+1; s<sublastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            ApplicantRelatedCompany applicantRelatedCompany = new ApplicantRelatedCompany();//申请人关联企业
                            //关联企业名称
                            applicantRelatedCompany.setRelatedCompanyName(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)));
                            //是否参与经营.  加3列
                            String isInvolvedInOperation = getStringValue(sheet.getRow(s).getCell(sublastColumn+4));
                            if (StringUtils.isNotEmpty(isInvolvedInOperation)){
                                applicantRelatedCompany.setIsInvolvedInOperation(Boolean.valueOf(isInvolvedInOperation));
                            }
                            //股份占比(%)。  加5列
                            String stockholdingRatio = getStringValue(sheet.getRow(s).getCell(sublastColumn+6));
                            if (StringUtils.isNotEmpty(stockholdingRatio)){
                                applicantRelatedCompany.setStockholdingRatio(Double.valueOf(stockholdingRatio));
                            }
                            //加入集合
                            applicantRelatedCompanyList.add(applicantRelatedCompany);
                        }
                    }
                    //设置申请人关联企业
                    electronicSurveyReport.setApplicantRelatedCompanies(applicantRelatedCompanyList);
                }
                //申请人及关联主体银行融资及对外担保（最多维护6条）单独处理。 注意ElectronicSurveyReport不要对这个字段加注解
                if ("申请人及关联主体银行融资及对外担保（最多维护6条）".equals(getStringValue(row.getCell(j)))){
                    //开始判断“申请人及关联主体银行融资及对外担保（最多维护6条）占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<ApplicantBankFinanceAndGuarantee> applicantBankFinanceAndGuaranteeList = new ArrayList<>();
                    int subfirstRow = 0;
                    int sublastRow = 0;
                    int sublastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        subfirstRow = cellRangeAddress.getFirstRow();
                        sublastRow = cellRangeAddress.getLastRow();
                        sublastColumn = cellRangeAddress.getLastColumn();
                        for (int s= subfirstRow+1; s<sublastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            ApplicantBankFinanceAndGuarantee applicantBankFinanceAndGuarantee = new ApplicantBankFinanceAndGuarantee();//申请人及关联主体银行融资及对外担保（最多维护6条）
                            //借款人名称
                            applicantBankFinanceAndGuarantee.setBorrowerName(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)));
                            //授信银行名称
                            applicantBankFinanceAndGuarantee.setCreditBankName(getStringValue(sheet.getRow(s).getCell(sublastColumn+2)));
                            //抵质押类授信敞口余额(万元)
                            String collateralLoanBalance = getStringValue(sheet.getRow(s).getCell(sublastColumn+3));
                            if (StringUtils.isNotEmpty(collateralLoanBalance)){
                                applicantBankFinanceAndGuarantee.setCollateralLoanBalance(Double.valueOf(collateralLoanBalance));
                            }
                            //保证类授信敞口余额(万元)
                            String guaranteeLoanBalance = getStringValue(sheet.getRow(s).getCell(sublastColumn+4));
                            if (StringUtils.isNotEmpty(guaranteeLoanBalance)){
                                applicantBankFinanceAndGuarantee.setGuaranteeLoanBalance(Double.valueOf(guaranteeLoanBalance));
                            }
                            //联保类授信敞口余额(万元)
                            String jointGuaranteeLoanBalance = getStringValue(sheet.getRow(s).getCell(sublastColumn+5));
                            if (StringUtils.isNotEmpty(jointGuaranteeLoanBalance)){
                                applicantBankFinanceAndGuarantee.setJointGuaranteeLoanBalance(Double.valueOf(jointGuaranteeLoanBalance));
                            }
                            //其他授信敞口余额(万元)
                            String otherLoanBalance = getStringValue(sheet.getRow(s).getCell(sublastColumn+6));
                            if (StringUtils.isNotEmpty(otherLoanBalance)){
                                applicantBankFinanceAndGuarantee.setOtherLoanBalance(Double.valueOf(otherLoanBalance));
                            }
                            //是否本次置换
                            String isReplacementThisTime = getStringValue(sheet.getRow(s).getCell(sublastColumn+7));
                            if (StringUtils.isNotEmpty(isReplacementThisTime)){
                                applicantBankFinanceAndGuarantee.setIsReplacementThisTime(Boolean.valueOf(isReplacementThisTime));
                            }
                            //加入集合
                            applicantBankFinanceAndGuaranteeList.add(applicantBankFinanceAndGuarantee);
                        }
                    }
                    //设置申请人及关联主体银行融资及对外担保（最多维护6条）
                    electronicSurveyReport.setAndGuarantees(applicantBankFinanceAndGuaranteeList);
                }
                //借款人及关联主体对外担保明细（最多维护3条）单独处理。 注意ElectronicSurveyReport不要对这个字段加注解
                if ("借款人及关联主体对外担保明细（最多维护3条）".equals(getStringValue(row.getCell(j)))){
                    //开始判断“借款人及关联主体对外担保明细（最多维护3条）占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<GuarantorDetail> guarantorDetailList = new ArrayList<>();
                    int subfirstRow = 0;
                    int sublastRow = 0;
                    int sublastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        subfirstRow = cellRangeAddress.getFirstRow();
                        sublastRow = cellRangeAddress.getLastRow();
                        sublastColumn = cellRangeAddress.getLastColumn();
                        for (int s= subfirstRow+1; s<sublastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            GuarantorDetail guarantorDetail = new GuarantorDetail();//借款人及关联主体对外担保明细（最多维护3条）
                            //保证人名称
                            guarantorDetail.setGuarantorName(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)));
                            //被保证人名称
                            guarantorDetail.setGuaranteedName(getStringValue(sheet.getRow(s).getCell(sublastColumn+2)));
                            //担保金额（万元）
                            String guaranteeAmount = getStringValue(sheet.getRow(s).getCell(sublastColumn+3));
                            if (StringUtils.isNotEmpty(guaranteeAmount)){
                                guarantorDetail.setGuaranteeAmount(Double.valueOf(guaranteeAmount));
                            }
                            //被保证人授信银行名称
                            guarantorDetail.setBankNameOfGuaranteed(getStringValue(sheet.getRow(s).getCell(sublastColumn+4)));
                            //担保对应的授信金额（万元）
                            String correspondingCreditLine = getStringValue(sheet.getRow(s).getCell(sublastColumn+5));
                            if (StringUtils.isNotEmpty(correspondingCreditLine)){
                                guarantorDetail.setCorrespondingCreditLine(Double.valueOf(correspondingCreditLine));
                            }
                            //被保证人与保证人关系
                            guarantorDetail.setRelationshipBetweenParties(getStringValue(sheet.getRow(s).getCell(sublastColumn+6)));
                            //是否为联保
                            String isJointGuarantee = getStringValue(sheet.getRow(s).getCell(sublastColumn+7));
                            if (StringUtils.isNotEmpty(isJointGuarantee)){
                                guarantorDetail.setIsJointGuarantee(Boolean.valueOf(isJointGuarantee));
                            }
                            //加入集合
                            guarantorDetailList.add(guarantorDetail);
                        }
                    }
                    //设置借款人及关联主体对外担保明细（最多维护3条）
                    electronicSurveyReport.setGuarantorDetails(guarantorDetailList);
                }
                //申请人家庭资产（针对个人经营者，最多维护3条）单独处理。 注意ElectronicSurveyReport不要对这个字段加注解
                if ("申请人家庭资产（针对个人经营者，最多维护3条）".equals(getStringValue(row.getCell(j)))){
                    //开始判断“申请人家庭资产（针对个人经营者，最多维护3条）占据的行数。
                    List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                    CellRangeAddress cellRangeAddress = getCellRangeAddress(mergedRegions, row.getCell(j));
                    List<ApplicantFamilyAsset> applicantFamilyAssetList = new ArrayList<>();
                    int subfirstRow = 0;
                    int sublastRow = 0;
                    int sublastColumn = 0;
                    if (Objects.nonNull(cellRangeAddress)) {
                        subfirstRow = cellRangeAddress.getFirstRow();
                        sublastRow = cellRangeAddress.getLastRow();
                        sublastColumn = cellRangeAddress.getLastColumn();
                        for (int s= subfirstRow+1; s<sublastRow; s++){
                            if (StringUtils.isEmpty(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)))){
                                continue;//无值不处理，只判断第一个
                            }
                            ApplicantFamilyAsset applicantFamilyAsset = new ApplicantFamilyAsset();//申请人家庭资产（针对个人经营者，最多维护3条）
                            //坐落
                            applicantFamilyAsset.setLocation(getStringValue(sheet.getRow(s).getCell(sublastColumn+1)));
                            //市值（万元）
                            String marketValue = getStringValue(sheet.getRow(s).getCell(sublastColumn+2));
                            if (StringUtils.isNotEmpty(marketValue)){
                                applicantFamilyAsset.setMarketValue(Double.valueOf(marketValue));
                            }
                            //所有权人
                            applicantFamilyAsset.setOwner(getStringValue(sheet.getRow(s).getCell(sublastColumn+3)));
                            //共有人
                            applicantFamilyAsset.setCoOwner(getStringValue(sheet.getRow(s).getCell(sublastColumn+4)));
                            //房屋（土地）类型
                            applicantFamilyAsset.setPropertyType(getStringValue(sheet.getRow(s).getCell(sublastColumn+5)));
                            //房屋（土地）面积（㎡）
                            String propertyArea = getStringValue(sheet.getRow(s).getCell(sublastColumn+6));
                            if (StringUtils.isNotEmpty(propertyArea)){
                                applicantFamilyAsset.setPropertyArea(Double.valueOf(propertyArea));
                            }
                            //抵押状况
                            applicantFamilyAsset.setMortgageStatus(getStringValue(sheet.getRow(s).getCell(sublastColumn+7)));
                            //出租状况
                            applicantFamilyAsset.setLeaseStatus(getStringValue(sheet.getRow(s).getCell(sublastColumn+8)));
                            //加入集合
                            applicantFamilyAssetList.add(applicantFamilyAsset);
                        }
                    }
                    //设置借申请人家庭资产（针对个人经营者，最多维护3条）
                    electronicSurveyReport.setApplicantFamilyAssets(applicantFamilyAssetList);
                }
            }
        }
        //将map值赋值到申请人对象中去
        if (!map.isEmpty()){
            setMapParam(electronicSurveyReport, map);
        }
        return electronicSurveyReport;
    }
    /**
     * 解析“签署意见”。 特殊处理
     * @param firstRow 起始行
     * @param lastRow 结束行
     * @param sheet
     * @return String
     */
    private static String analysisExcelOpinion(int firstRow, int lastRow, Sheet sheet){
        try {
            return getStringValue(sheet.getRow(firstRow).getCell(2));//直接取第3列值
        }catch (Exception e){
            System.out.println("解析签署意见出错："+e.getMessage());
        }
        return "";
    }

    /**
     * 判断单元格是否在合并单元格内
     * @param mergedRegions 所有合并单元格集合
     * @param cell 校验单元格
     * @return
     */
    private static CellRangeAddress getCellRangeAddress(List<CellRangeAddress> mergedRegions, Cell cell) {
        if (CollectionUtil.isEmpty(mergedRegions)) {
            return null;
        }
        if (Objects.isNull(cell)) {
            return null;
        }
        for (CellRangeAddress mergedRegion : mergedRegions) {
            Boolean in = mergedRegion.isInRange(cell);
            if (in) {
                return mergedRegion;
            }
        }
        return null;
    }

    /**
     * 获取单元格值
     * @param cellRangeAddress 合并单元格
     * @param sheet excel页面
     * @return
     */
    private static String getCellRangeAddressValue(CellRangeAddress cellRangeAddress, Sheet sheet) {
        if (Objects.isNull(cellRangeAddress)) {
            return StringUtils.EMPTY;
        }
        if (Objects.isNull(sheet)) {
            return StringUtils.EMPTY;
        }
        Row row = sheet.getRow(cellRangeAddress.getFirstRow());
        if (Objects.isNull(row)) {
            return StringUtils.EMPTY;
        }
        Cell cell = row.getCell(cellRangeAddress.getFirstColumn());
        if (Objects.isNull(cell)) {
            return StringUtils.EMPTY;
        }
        return getStringValue(cell);
    }

    /**
     * 获取单元格值
     * @param cell 单元格
     * @return
     */
    private static String getStringValue(Cell cell) {
        String value = "";
        // 判断单元格类型
        if (cell.getCellType() == CellType.STRING) {
            value = cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            value = String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == CellType.NUMERIC) {
            value = String.valueOf(cell.getNumericCellValue());
        }
        return value;
    }

    /**
     * 将map值存入body中
     * @param body
     * @param map
     */
    public static void setMapParam(Object body, Map<String,Object> map){
        if (null == body || null == map || map.isEmpty()){
            return;
        }
        Class reqClass = body.getClass();
        Method[] methods = reqClass.getMethods();
        for (Method method:methods){
            try {
                String getName = method.getName();
                if (!getName.startsWith("get") || "getClass".equals(getName)){
                    continue;
                }
                //获取去除get后的字符串，例如 getName---> Name
                String field = getName.substring(3);
                //首字母转小写，例如 Nmae--->name
                field = field.substring(0,1).toLowerCase() + field.substring(1);
                //map中不包含这个key，跳过不处理
                if (!map.containsKey(field)){
                    continue;
                }
                //拼接set方法，将首字母g换成s。 例如：getName--->setName
                String meSetName = "s"+getName.substring(1);
                //获取get方法返回值，和set方法入参一致
                Class returnClass = method.getReturnType();
                //调用set方法赋值
                Method setMethod = reqClass.getMethod(meSetName, returnClass);
                Object value = map.get(field);
                //空值处理, 截至当前value存的都是string
                if (StringUtils.isEmpty(value.toString())){
                    continue;
                }
                //部分类型特殊处理
                String simpleName = method.getReturnType().getSimpleName();
                if ("BigDecimal".equals(simpleName)){
                    value = new BigDecimal(value.toString());
                }else if ("Integer".equals(simpleName)){
                    value = Integer.valueOf(value.toString());
                }else if ("Boolean".equals(simpleName)){
                    value = Boolean.valueOf(value.toString());
                }else if ("Double".equals(simpleName)){
                    value = Double.valueOf(value.toString());
                }
                //对象set值
                setMethod.invoke(body, value);
            }catch (Exception e){
                System.out.println("赋值异常，异常原因：" + reqClass.getName() + "--"+method.getName()+";" + e.getMessage());
            }
        }
    }

    /**
     * 判断单元格值是否能映射指定类的‘注解’
     * 是：返回类变量
     * 否：返回空
     * @param object 模块对象
     * @param cell
     * @return
     */
    private static String checkExcelAttribute(Object object, Cell cell){
        try {
            String cellValue = getStringValue(cell);
            if (StringUtils.isEmpty(cellValue)){
                return "";
            }
            //获取类的变量
            Field[] fields = object.getClass().getDeclaredFields();
            for (Field field:fields){
                //如果变量没有注解‘Excel’，跳过该字段
                if (Objects.isNull(field.getDeclaredAnnotation(Excel.class))){
                    continue;
                }
                //注解名称值，name
                String reqField = field.getDeclaredAnnotation(Excel.class).name();
                //如果注解值等于cell单元格的值，则返回对应的变量值
                if (reqField.equals(getStringValue(cell))){
                    return field.getName();//实际变量名
                }
            }
        }catch (Exception e){
            System.out.println("解析调查报告模块对象注解时报错，msg："+e.getMessage());
        }
        return "";
    }
}
