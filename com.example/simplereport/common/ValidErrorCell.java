package com.example.simplereport.common;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

@Data
public class ValidErrorCell {
    @Excel(name="错误选项名称",orderNum = "0")
    private String columnName;
    @Excel(name = "错误信息",orderNum = "1")
    private String errorMsg;
}
