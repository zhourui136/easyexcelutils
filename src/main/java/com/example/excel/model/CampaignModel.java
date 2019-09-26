package com.example.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import com.feinik.excel.annotation.ExcelValueFormat;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * @author zhourui
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class CampaignModel extends BaseRowModel implements Serializable {

    @ExcelProperty(value = {"代扣款项1", "旷罚扣款"}, index = 0)
    private String day1;
    @ExcelProperty(value = {"代扣款项1", "请假扣款"}, index = 1)
    private String day2;
    @ExcelProperty(value = {"代扣款项1", "入离职缺勤扣款"}, index = 2)
    private String day3;
    @ExcelProperty(value = {"代扣款项1", "其他"}, index = 3)
    private String day4;

    @ExcelProperty(value = {"本月应发","本月应发"}, index = 4)
    private String campaignId;

    @ExcelProperty(value = {"代扣款项2", "本月社保"}, index = 5)
    private String campaignName;
    @ExcelProperty(value = {"代扣款项2", "社保缴费"}, index = 6)
    private String campaignName1;

    @ExcelProperty(value = "", index = 7)
    @ExcelValueFormat(format = "{0}$")
    private String cost;

    @ExcelProperty(value = "点击次数", index = 8)
    private String clicks;

    @ExcelProperty(value = "点击率", index = 9)
    @ExcelValueFormat(format = "{0}%")
    private String ctr;

}
