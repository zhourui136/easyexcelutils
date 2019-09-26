package com.example.excel;

import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.example.excel.handler.UserDataHandler;
import com.example.excel.model.User;
import com.example.excel.utils.ExcelUtils;
import com.feinik.excel.handler.ExcelDataHandler;
import com.feinik.excel.utils.ExcelUtil;
import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
@Slf4j
public class ExcelApplicationTests {

    @Value("${file.path}")
    private String filePath;

    @Test
    public void test() throws Exception {
        List<User> userList = Lists.newArrayList();
        User user1 = new User().setUserId(1001).setUserName("张三").setScore(80);
        User user2 = new User().setUserId(1002).setUserName("李四").setScore(90);
        userList.add(user1);
        userList.add(user2);
        File file = new File(filePath);
        if (!file.exists()) {
            file.mkdirs();
        }
        String fileName = filePath + "/" + "用户表" + ExcelTypeEnum.XLSX.getValue();
        File excel = new File(fileName);
        if (!excel.exists()) {
            excel.createNewFile();
        }
        //ExcelUtil.writeExcelWithOneSheet(new File(fileName),"user",true, userList, new UserDataHandler());
        Long startTime=System.currentTimeMillis();
        Boolean flag = ExcelUtils.writeExcelWithOneSheet(new FileOutputStream(excel), "用户表", true, userList, new UserDataHandler());
        Long endTime=System.currentTimeMillis();
        log.info("文件导出用时:"+(endTime-startTime)+"毫秒");
        //ExcelUtils.writeExcelWithOneSheet(new FileOutputStream(excel), "用户表", false, userList, new UserDataHandler());
    }
}
