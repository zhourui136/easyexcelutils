package com.example.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * @author zhourui
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@NoArgsConstructor
@AllArgsConstructor
public class User extends BaseRowModel {
    @ExcelProperty(value = {"用户编号", "用户名", "用户年龄"}, index = 0)
    private Integer userId;

    @ExcelProperty(value = "用户姓名", index = 1)
    private String userName;

    @ExcelProperty(value = "所得分", index = 2)
    private Integer score;
}
