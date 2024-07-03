package com.example.simplereport.common;

import com.example.simplereport.entity.*;

import java.util.List;

public interface BusinessConstant {
    /**
     * 定义了类集合，这里是简单报表中各个实体对象
     *
     * 名称后期优化
     */
    class ClassPath{
        public static Class SURVEY_SIMPLE_REPORT_ONE = ApplicantBaseInfo.class;
        public static Class SURVEY_SIMPLE_REPORT_TWO = ApplicantBusinessInfo.class;
        public static Class SURVEY_SIMPLE_REPORT_THREE = CollateralInfo.class;
        public static Class SURVEY_SIMPLE_REPORT_FOUR = ApplicantReportBaseInfo.class;
        public static Class SURVEY_SIMPLE_REPORT_FIVE = ApplicantBankFinanceAndGuarantee.class;
        public static Class SURVEY_SIMPLE_REPORT_SIX = ApplicantReportCreditAnalysisConclusion.class;
        public static Class SURVEY_SIMPLE_REPORT_SEVEN = MortgageAssetOtherInfo.class;
        public static Class SURVEY_SIMPLE_REPORT_EIGHT = BusinessCriticalIndicator.class;
        public static Class SURVEY_SIMPLE_REPORT_NINE = CreditInquiryParent.class;
        public static Class SURVEY_SIMPLE_REPORT_TEN = CreditExposureCalculationNew.class;
        public static Class SURVEY_SIMPLE_REPORT_ELEVEN = CreditExposureCalculationNew.class;
        public static Class SURVEY_SIMPLE_REPORT_TWELVE = ElectronicSurveyReport.class;
        public static Class LIST_CLASS_REFERENCE_PATH = List.class;

        public static Class[] SURVEY_SIMPLE_REPORT_FIELD = {SURVEY_SIMPLE_REPORT_ONE,SURVEY_SIMPLE_REPORT_TWO,SURVEY_SIMPLE_REPORT_THREE,
                SURVEY_SIMPLE_REPORT_FOUR,SURVEY_SIMPLE_REPORT_FIVE,SURVEY_SIMPLE_REPORT_SIX,SURVEY_SIMPLE_REPORT_SEVEN,
                SURVEY_SIMPLE_REPORT_EIGHT,SURVEY_SIMPLE_REPORT_NINE,SURVEY_SIMPLE_REPORT_TEN,SURVEY_SIMPLE_REPORT_ELEVEN,
                SURVEY_SIMPLE_REPORT_TWELVE};
    }

    /**
     * 纵向遍历-列单元格合并不规律集合
     *
     * 对应实体中集合属性名称
     */
    class CollectionName{
        public static String NOT_RULE_OBJECT_ONE = "appraisalInformations";
        public static String[] NOT_RULE_OBJECT_ARRAY={NOT_RULE_OBJECT_ONE};
    }
}
