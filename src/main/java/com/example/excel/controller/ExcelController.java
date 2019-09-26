package com.example.excel.controller;

import com.alibaba.excel.metadata.BaseRowModel;
import com.example.excel.handler.CampaignDataHandler;
import com.example.excel.model.CampaignModel;
import com.example.excel.utils.ExcelUtils;
import com.google.common.collect.Lists;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author zhourui
 */
@RestController
@RequestMapping("/excel")
public class ExcelController {
    @Value("${file.path}")
    private String filePath;

    @RequestMapping("/exportExcel")
    public void exportExcel() throws Exception {
        CampaignModel m1 = new CampaignModel("2019-01-01", "1", "1", "1", "1", "10000000", "campaign1", "12.21", "100", "0.11");
        CampaignModel m2 = new CampaignModel("2019-01-02", "1", "1", "1", "1", "12001010", "campaign2", "13", "99", "0.91");
        CampaignModel m3 = new CampaignModel("2019-01-03", "1", "1", "1", "1", "12001010", "campaign3", "10", "210", "1.13");
        CampaignModel m4 = new CampaignModel("2019-01-04", "1", "1", "1", "1", "15005010", "campaign4", "21.9", "150", "0.15");

        ArrayList<CampaignModel> data1 = Lists.newArrayList(m1, m2);
        ArrayList<CampaignModel> data2 = Lists.newArrayList(m3, m4);
        Map<String, List<? extends BaseRowModel>> map = new HashMap<>();
        map.put("sheet1", data1);
        map.put("sheet2", data2);
        File file = new File(filePath);
        if (!file.exists()) {
            file.mkdirs();
        }
        String fileName = filePath + "/工资表.xlsx";
        File excel = new File(fileName);
        if (!excel.exists()) {
            excel.createNewFile();
        }
       // ExcelUtils.writeExcelWithMultiSheet(new FileOutputStream(excel), map, true, new CampaignDataHandler());
    }
}
