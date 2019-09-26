package com.example.excel.handler;

import com.feinik.excel.handler.ExcelDataHandler;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * @author zhourui
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class UserDataHandler implements ExcelDataHandler {

    /**
     * 通过对象池的方式来解决大数据量下重复创建CellStyle而导致的异常
     * 相同设置属性的CellStyle对象达到可复用
     */
    private Map<String, CellStyle> stylePool = new HashMap<>();

    @Override
    public void headFont(Font font, int cellIndex) {
        font.setColor(IndexedColors.WHITE.getIndex());
    }

    @Override
    public void headCellStyle(CellStyle cellStyle, int cellIndex) {
        if(cellIndex==1){
            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        }
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
    }

    @Override
    public void contentFont(Font font, int i, Object o) {
        font.setColor(IndexedColors.WHITE.getIndex());
    }

    @Override
    public void contentCellStyle(CellStyle cellStyle, int i) {
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
    }

    @Override
    public void sheet(int i, Sheet sheet) {
    }
}
