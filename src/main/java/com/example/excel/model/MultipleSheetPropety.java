package com.example.excel.model;

import com.alibaba.excel.metadata.Sheet;
import lombok.Data;

import java.util.List;

/**
 * @author zhourui
 */
@Data
public class MultipleSheetPropety {

    private List<List<Object>> data;

    private Sheet sheet;
}
