package com.example.simplereport.common;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.hutool.core.util.ObjectUtil;
import com.alibaba.fastjson.JSONArray;
import com.example.simplereport.entity.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;
import static com.example.simplereport.common.BusinessConstant.*;

public class CommonUtil {
    /**
     * 模板对象赋值--正常从service获取
     *
     * @return
     */
    public static SurveySimpleReport getSurveySimpleReport() {
        SurveySimpleReport simpleReport = new SurveySimpleReport();

        //申请人进件基础信息
        ApplicantBaseInfo baseInfo = new ApplicantBaseInfo();
        baseInfo.setApplicantIdNumber("410223199306278976");
        baseInfo.setApplicantName("汉娜");
        baseInfo.setApplicantPhoneNumber("18189705672");
        baseInfo.setCreditCode("9080765423267");
        baseInfo.setEnterpriseName("阿里巴巴");
        baseInfo.setSpouseIdNumber("411228199506235894");
        baseInfo.setSpouseName("里斯");
        baseInfo.setSpousePhoneNumber("189765423456");
        baseInfo.setLoanApplicant("借款主体_选择1");
        simpleReport.setApplicantBaseInfo(baseInfo);

        //申请人业务信息
        ApplicantBusinessInfo applicantBusinessInfo = new ApplicantBusinessInfo();
        applicantBusinessInfo.setWithdrawalMethod("提款方式_1");
        applicantBusinessInfo.setBusinessVariety("业务品种_1");
        applicantBusinessInfo.setBusinessType("业务类型_1");
        applicantBusinessInfo.setApplicationAmount(new BigDecimal("75896.562"));
        applicantBusinessInfo.setPaymentMethod("支付宝");
        applicantBusinessInfo.setTermMonths(10);
        applicantBusinessInfo.setSingleNoteMaxTermMonths(10);
        applicantBusinessInfo.setOverduePenaltyRateIncrease(new BigDecimal("59.63"));
        applicantBusinessInfo.setAnnualInterestRate(new BigDecimal("0.75"));
        applicantBusinessInfo.setPurpose("用途_测试");
        //小企业特色业务
        applicantBusinessInfo.setSmallBusinessFeatures1("特色业务_1");
        applicantBusinessInfo.setSmallBusinessFeatures2("特色业务_2");
        applicantBusinessInfo.setSmallBusinessFeatures3("特色业务_3");
        applicantBusinessInfo.setSmallBusinessFeatures4("特色业务_4");
        applicantBusinessInfo.setSmallBusinessFeatures5("特色业务_5");
        //主要业务
        applicantBusinessInfo.setMainProduct("特色业务_1");

        applicantBusinessInfo.setIndustryOrientation1("陕西省");
        applicantBusinessInfo.setIndustryOrientation2("西安市");
        applicantBusinessInfo.setIndustryOrientation3("小西安市");
        applicantBusinessInfo.setIndustryOrientation4("未央区");
        simpleReport.setApplicantBusinessInfo(applicantBusinessInfo);


        //申请人家庭资产
        List<ApplicantFamilyAsset> applicantFamilyAssets = new ArrayList<>();
        ApplicantFamilyAsset applicantFamilyAsset = new ApplicantFamilyAsset();
        applicantFamilyAsset.setCoOwner("1");
        applicantFamilyAsset.setLeaseStatus("2");
        applicantFamilyAsset.setLocation("3");
        applicantFamilyAsset.setMarketValue(new Double("500"));
        applicantFamilyAsset.setMortgageStatus("抵押状况_2");
        applicantFamilyAsset.setLeaseStatus("出租状况_3");
        applicantFamilyAssets.add(applicantFamilyAsset);

        ElectronicSurveyReport surveyReport = new ElectronicSurveyReport();
        surveyReport.setApplicantFamilyAssets(applicantFamilyAssets);
        simpleReport.setElectronicSurveyReport(surveyReport);

        // 评估信息
        AppraisalInformation information = new AppraisalInformation();
        information.setAppraisalSource("测试");
        information.setAppraisalTotalPrice(new BigDecimal("5689.12"));
        information.setAppraisalUnitPrice(new BigDecimal("89574.12"));

        AppraisalInformation information1 = new AppraisalInformation();
        information1.setAppraisalSource("测试1");
        information1.setAppraisalTotalPrice(new BigDecimal("56891.12"));
        information1.setAppraisalUnitPrice(new BigDecimal("895741.12"));

        ArrayList<AppraisalInformation> informationList = new ArrayList<>();
        informationList.add(information);
        informationList.add(information1);
        MortgageAssetOtherInfo mortgageAssetOtherInfo = new MortgageAssetOtherInfo();
        mortgageAssetOtherInfo.setAppraisalInformations(informationList);

        simpleReport.setAssetOtherInfo(mortgageAssetOtherInfo);
        return simpleReport;
    }

    /**
     * 定义块状对象容器
     *
     * ObjectName对应得是块状对象在父对象中字段名称
     * 不规则块状对象注意在BusinessConstant.CollectionName中预定义
     * ObjectName是块状对象容器，下拉框容易，与数据集关联得唯一值
     * @return
     */
    public static List<BlockAreaContainer> getBlockObjContainerList() {
        List<BlockAreaContainer> containers = new ArrayList<>();

        BlockAreaContainer container = new BlockAreaContainer();
        container.setBaseRowIndex(87);
        container.setBaseColIndex(2);
        container.setMergeColIndex(1);
        container.setObjectName("applicantFamilyAssets");
        containers.add(container);

        //评估信息
        BlockAreaContainer container1 = new BlockAreaContainer();
        container1.setBaseRowIndex(34);
        container1.setBaseColIndex(2);
        container1.setObjectName(CollectionName.NOT_RULE_OBJECT_ONE);
        containers.add(container1);
        return containers;
    }
    /**
     * 这里获取各个下拉框的预设值
     * redis 存储时以对应属性名为key存储，value为预设值，后期值从数据库中获取
     * fieldName是类名称_字段名称，
     * @return
     */
    public static List<ExcelSelectorContainer> getSelectorDto() {
        List<ExcelSelectorContainer> selectorDtos = new ArrayList<>();
        //借款主体
        String[] loanApplicantDatas = {"借款主体_选择1", "借款主体_选择2"};
        selectorDtos.add(setSelector("applicantBaseInfo_loanApplicant",1,2, loanApplicantDatas,false));

        //业务品种
        String[] businessVarietys = {"业务品种_1", "业务品种_2"};
        selectorDtos.add(setSelector("applicantBusinessInfo_businessVariety",4,2, businessVarietys,false));

        //业务类型
        String[] businessTypes = {"业务类型_1", "业务类型_2"};
        selectorDtos.add(setSelector("applicantBusinessInfo_businessVariety",4,4, businessTypes,false));

        //支付方式
        String[] paymentMethods = {"支付宝", "微信"};
        selectorDtos.add(setSelector("applicantBusinessInfo_paymentMethod",5,2, paymentMethods,false));

        //小企业特色业务（可多选）
        String[] smallBusinessFeaturess = {"特色业务_1", "特色业务_2", "特色业务_3", "特色业务_4", "特色业务_5"};
        selectorDtos.add(setSelector("applicantBusinessInfo_smallBusinessFeatures1",8,2, smallBusinessFeaturess,false));
        selectorDtos.add(setSelector("applicantBusinessInfo_smallBusinessFeatures2",8,3, smallBusinessFeaturess,false));
        selectorDtos.add(setSelector("applicantBusinessInfo_smallBusinessFeatures3",8,4, smallBusinessFeaturess,false));
        selectorDtos.add(setSelector("applicantBusinessInfo_smallBusinessFeatures4",8,5, smallBusinessFeaturess,false));
        selectorDtos.add(setSelector("applicantBusinessInfo_smallBusinessFeatures5",8,6, smallBusinessFeaturess,false));
        selectorDtos.add(setSelector("applicantBusinessInfo_mainProduct",8,8, smallBusinessFeaturess,false));

        /**
         * 遍历下拉
         */
        String[] mortgageStatusArray = {"抵押状况_1", "抵押状况_2", "抵押状况_3", "抵押状况_4", "抵押状况_5"};
        selectorDtos.add(setSelector("applicantFamilyAssets_mortgageStatus",0,0,  mortgageStatusArray,true));
        String[] leaseStatusArray = {"出租状况_1", "出租状况_2", "出租状况_3", "出租状况_4", "出租状况_5"};
        selectorDtos.add(setSelector("applicantFamilyAssets_leaseStatus",0,0,leaseStatusArray,true));


        return selectorDtos;
    }
    public static List<MultiLevelDropDownContainer> getMultiLevelDropDownContainers() {
        List<MultiLevelDropDownContainer> containers = new ArrayList<>();
        MultiLevelDropDownContainer container = new MultiLevelDropDownContainer();
        container.setFirstRow(9);
        container.setLastRow(9);
        container.setFirstCol(2);
        container.setLastCol(2);
        // container.setMergeNum(2);
        container.setLevel(1);
        //后期级联工具使用
        //container.setData(new HashMap());
        container.setFieldName("applicantBusinessInfo_industryOrientation1");
        containers.add(container);

        MultiLevelDropDownContainer container1 = new MultiLevelDropDownContainer();
        container1.setFirstRow(9);
        container1.setLastRow(9);
        container1.setFirstCol(4);
        container1.setLastCol(4);
        container1.setLevel(2);
        //后期级联工具使用
        //container.setData(new HashMap());
        container1.setFieldName("applicantBusinessInfo_industryOrientation2");
        containers.add(container1);

        return containers;
    }

    /**
     * 写入对象
     * @param fieldName
     * @param dataArray
     * @return
     */
    public static ExcelSelectorContainer setSelector(String fieldName, int firstRow, int firstCol, String[] dataArray, boolean isForeach) {
        return setSelector(fieldName,firstRow,firstRow,firstCol,firstCol,dataArray,isForeach);
    }


    /**
     * 下拉框实体写入
     *
     * @param fieldName
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     * @param dataArray
     * @return
     */
    public static ExcelSelectorContainer setSelector(String fieldName, int firstRow, int lastRow, int firstCol, int lastCol, String[] dataArray, boolean isForeach) {
        ExcelSelectorContainer selectorDto1 = new ExcelSelectorContainer();
        selectorDto1.setFieldName(fieldName);
        selectorDto1.setFirstRow(firstRow);
        selectorDto1.setLastRow(lastRow);
        selectorDto1.setFirstCol(firstCol);
        selectorDto1.setLastCol(lastCol);

        selectorDto1.setDatas(dataArray);

        selectorDto1.setForForeach(isForeach);
        return selectorDto1;
    }

    public static TreeMap<String, TreeMap<String, TreeMap<String, List<String>>>> dropDownDataSource(){
        TreeMap<String, TreeMap<String, TreeMap<String, List<String>>>> selectTree = new TreeMap<>();

        TreeMap<String, List<String>> map_henan = new TreeMap<>();
        map_henan.put( "小郑州市", JSONArray.parseArray( "[ \"二七区\",\"登封市\",\"新郑市\" ]",String.class )  );
        map_henan.put( "小洛阳市",JSONArray.parseArray( "[ \"洛龙区\",\"涧西区\" ]",String.class ) );

        TreeMap<String, List<String>> map_shanxi = new TreeMap<>();
        map_shanxi.put( "小西安市",JSONArray.parseArray( "[ \"未央区\",\"长安区\",\"高陵区\" ]",String.class ) );

        TreeMap<String, TreeMap<String, List<String>>> map_henan1 = new TreeMap<>();
        map_henan1.put("郑州市",map_henan);
        TreeMap<String, TreeMap<String, List<String>>> map_shanxi1 = new TreeMap<>();
        map_shanxi1.put("西安市",map_shanxi);


        selectTree.put( "河南省",map_henan1 );
        selectTree.put( "山西省",map_shanxi1 );
        return selectTree;
    }


    /**
         * 递归获取所有属性名与属性值
         *
         * @param obj 当前要解析的对象
         * @return 所有属性名与属性值的列表
         */
    public static List<String[]> getAllFieldsAndValues(Object obj) throws IllegalAccessException {
        List<String[]> fieldsAndValues = new ArrayList<>();
        if (obj == null) return fieldsAndValues;

        Class<?> clazz = obj.getClass();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true); // 允许访问私有字段

            Object fieldValue = field.get(obj);
            if (fieldValue != null) {
                // 如果字段是一个list集合
                if(ClassPath.LIST_CLASS_REFERENCE_PATH.equals(field.getType())){
                    continue;
                }
                // 如果字段是另一个实体类，则递归解析
                //if (!field.getType().isPrimitive() && !field.getType().equals(String.class)) {
                if (Arrays.asList(ClassPath.SURVEY_SIMPLE_REPORT_FIELD).contains(field.getType())) {
                    List<String[]> nestedFieldsAndValues = getAllFieldsAndValues(fieldValue);
                    for (String[] nestedPair : nestedFieldsAndValues) {
                        fieldsAndValues.add(new String[]{field.getName() + "_" + nestedPair[0], nestedPair[1]});
                    }
                } else {
                    fieldsAndValues.add(new String[]{field.getName(), String.valueOf(fieldValue)});
                }
            }
        }
        return fieldsAndValues;
    }

    /**
     * 获取对象中类型为List的属性
     * @param obj
     * @return
     * @throws IllegalAccessException
     */
    public static Map<String,List> getEntityList(Object obj,boolean isRule) throws IllegalAccessException {
        Map<String,List> entityListMap = new HashMap<>();
        if (obj == null) return entityListMap;

        Class<?> clazz = obj.getClass();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true); // 允许访问私有字段

            Object fieldValue = field.get(obj);
            if (fieldValue != null) {
                // 如果字段是一个list集合
                if (Arrays.asList(ClassPath.SURVEY_SIMPLE_REPORT_FIELD).contains(field.getType())) {
                    entityListMap.putAll(getEntityList(fieldValue,isRule));
                }else if(ClassPath.LIST_CLASS_REFERENCE_PATH.equals(field.getType())){
                    String fieldNme = field.getName();
                    if(!isRule && Arrays.asList(CollectionName.NOT_RULE_OBJECT_ARRAY).contains(fieldNme)){
                        entityListMap.put(fieldNme,Arrays.asList(fieldValue));
                    }else if(isRule && !Arrays.asList(CollectionName.NOT_RULE_OBJECT_ARRAY).contains(fieldNme)){
                        entityListMap.put(fieldNme,Arrays.asList(fieldValue));
                    }
                }
            }
        }
        return entityListMap;
    }
    public static List<ValidErrorCell> validTableColumn(List<ValidErrorCell> errorList,Object object){
        try {

            //获取类的变量
            Field[] fields = object.getClass().getDeclaredFields();
            for (Field field:fields){
                //如果变量没有注解‘Excel’，跳过该字段
                if (Objects.isNull(field.getDeclaredAnnotation(Excel.class))){
                    continue;
                }
                field.setAccessible(true); // 允许访问私有字段
                //注解名称值，name
                String reqField = field.getDeclaredAnnotation(Excel.class).name();
                Object fieldValue = field.get(object);

                if(ObjectUtil.isEmpty(fieldValue)){
                    ValidErrorCell errorColumn = new ValidErrorCell();
                    errorColumn.setColumnName(reqField);
                    errorColumn.setErrorMsg("值为空");
                    errorList.add(errorColumn);
                }

                if(reqField.contains("手机号")){
                    //todo 校验手机号
                    ValidErrorCell errorColumn = new ValidErrorCell();
                    errorColumn.setColumnName(reqField);
                    errorColumn.setErrorMsg("手机号异常！");
                    errorList.add(errorColumn);
                }
            }
        }catch (Exception e){
            System.out.println("解析调查报告模块对象注解时报错，msg："+e.getMessage());
        }
        return errorList;
    }


    public static File transferToFile(MultipartFile multipartFile) {
        File file = null;
        try {
            String originalFilename = multipartFile.getOriginalFilename();
            String[] filename = originalFilename.split("\\.");
            file = File.createTempFile(filename[0], filename[1]);
            multipartFile.transferTo(file);
            file.deleteOnExit();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return file;

    }
}
