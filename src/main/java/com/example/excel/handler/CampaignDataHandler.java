package com.example.excel.handler;

import com.example.excel.model.CampaignModel;
import com.feinik.excel.handler.ExcelDataHandler;
import org.apache.poi.ss.usermodel.*;

/**
 * @author zhourui
 */
public class CampaignDataHandler implements ExcelDataHandler {
    @Override
    public void headCellStyle(CellStyle style, int cellIndex) {
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
    }

    @Override
    public void headFont(Font font, int cellIndex) {
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontName("微软雅黑");
    }

    @Override
    public void contentCellStyle(CellStyle style, int cellIndex) {
        style.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    @Override
    public void sheet(int sheetIndex, Sheet sheet) {
        System.out.println("sheetIndex = [" + sheetIndex + "]");
    }

    @Override
    public void contentFont(Font font, int cellIndex, Object data) {
        CampaignModel campaign = (CampaignModel) data;
        switch (cellIndex) {
            case 4: //这里的值为Model对象中ExcelProperty注解里的index值
                if (Long.valueOf(campaign.getClicks()) > 100) { //表示将点击次数大于100的第4列也就是点击次数列的cell字体标记为红色
                    font.setColor(IndexedColors.RED.getIndex());
                    font.setFontName("微软雅黑");
                    font.setItalic(true);
                    font.setBold(true);
                }
                break;
        }
        font.setColor(IndexedColors.WHITE.getIndex());
    }
}
