package com.example.excel.utils;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Font;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.TableStyle;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.example.excel.model.MultipleSheetPropety;
import com.feinik.excel.annotation.ExcelValueFormat;
import com.feinik.excel.handler.ExcelDataHandler;
import com.feinik.excel.handler.StyleHandler;
import com.feinik.excel.utils.BeanUtils;
import com.sun.org.apache.xml.internal.resolver.readers.TR9401CatalogReader;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.util.CollectionUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.text.MessageFormat;
import java.util.*;

/**
 * @author zhourui
 */
@Slf4j
public class ExcelUtils {
    private static Sheet initSheet;

    static {
        //初始化sheet
        initSheet = new Sheet(1, 0);
        //设置自适应宽度
        initSheet.setAutoWidth(Boolean.TRUE);
    }

    /**
     * 模版生成excle
     *
     * @param fileOutputStream 文件流
     * @param data             数据
     * @param head             表头
     */
    public static void writeBySimple(FileOutputStream fileOutputStream, List<List<Object>> data, List<List<String>> head) {
        writeSimpleBySheet(fileOutputStream, data, head, null);
    }

    /**
     * 生成excel（单个sheet）
     *
     * @param fileOutputStream 输出文件流
     * @param data             数据
     * @param sheet            导出excel的sheet
     * @param head             表头
     */
    public static void writeSimpleBySheet(FileOutputStream fileOutputStream, List<List<Object>> data, List<List<String>> head, Sheet sheet) {

        sheet = (sheet != null) ? sheet : initSheet;

        //定义Excel正文背景颜色
        TableStyle tableStyle = new TableStyle();
        tableStyle.setTableContentBackGroundColor(IndexedColors.WHITE);

        tableStyle.setTableHeadBackGroundColor(IndexedColors.SKY_BLUE);


        Font font = new Font();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("宋体");

        Font headFont = new Font();
        headFont.setFontName("宋体");
        headFont.setFontHeightInPoints((short) 12);

        tableStyle.setTableContentFont(font);
        tableStyle.setTableHeadFont(headFont);

        if (head != null) {
            sheet.setHead(head);
            sheet.setTableStyle(tableStyle);
            sheet.setSheetName("考勤明细表");
        }
        Sheet sheet2 = new Sheet(1, 0);
        sheet2.setSheetName("考勤汇总表");
        sheet2.setHead(head);

        OutputStream outputStream = null;
        ExcelWriter writer = null;
        try {
            outputStream = fileOutputStream;
            writer = EasyExcelFactory.getWriter(outputStream);
            writer.write1(data, sheet);
        } catch (Exception e) {
            log.error("生成excel文件异常", e);
        } finally {
            try {
                if (writer != null) {
                    writer.finish();
                }

                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (Exception e) {
                log.error("excel文件导出失败, 失败原因：", e);
            }
        }

    }

    /**
     * 适用于简单表头
     * 生成excel （根据实体类导出）
     *
     * @param filePath 文件绝对路径
     * @param data     数据
     */
    public static void writeWithTemplate(String filePath, List<? extends BaseRowModel> data) {
        writeWithTemplateAndSheet(filePath, data, null);
    }

    /**
     * 生成多Sheet的excel
     *
     * @param filePath              绝对路径
     * @param multipleSheetPropetys sheet加数据的列表
     */
    public static void writeWithMultipleSheet(String filePath, List<MultipleSheetPropety> multipleSheetPropetys) {
        if (CollectionUtils.isEmpty(multipleSheetPropetys)) {
            return;
        }
        OutputStream outputStream = null;
        ExcelWriter writer = null;
        try {
            outputStream = new FileOutputStream(filePath);
            writer = EasyExcelFactory.getWriter(outputStream);
            for (MultipleSheetPropety multipleSheetPropety : multipleSheetPropetys) {
                Sheet sheet = multipleSheetPropety.getSheet() != null ? multipleSheetPropety.getSheet() : initSheet;
                writer.write1(multipleSheetPropety.getData(), sheet);
            }

        } catch (FileNotFoundException e) {
            log.error("找不到文件或文件路径错误, 文件：" + filePath);
        } finally {
            try {
                if (writer != null) {
                    writer.finish();
                }

                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                log.error("excel文件导出失败, 失败原因：", e);
            }
        }
    }

    /**
     * 生成多sheet的excel表，自定义excel表样式
     *
     * @param fileOutputStream 文件输出流
     * @param data             数据
     * @param needHead         是否需要表头
     * @param handler          样式处理
     * @return boolean
     * @throws Exception
     */
    public static boolean writeExcelWithMultiSheet(FileOutputStream fileOutputStream, Map<String, List<? extends BaseRowModel>> data, boolean needHead, ExcelDataHandler handler) throws Exception {
        StyleHandler sh = new StyleHandler(data);
        sh.setHandler(handler);
        return writeExcelWithHandler(fileOutputStream, data, needHead, sh);
    }

    /**
     * 导出单sheet的excel表（自定义样式）
     *
     * @param fileOutputStream 文件输出流
     * @param sheetName        sheet名
     * @param needHead         是否需要head
     * @param data             数据
     * @param handler          样式处理
     * @return boolean
     * @throws Exception
     */
    public static boolean writeExcelWithOneSheet(FileOutputStream fileOutputStream, String sheetName, boolean needHead, List<? extends BaseRowModel> data, ExcelDataHandler handler) throws Exception {
        Map<String, List<? extends BaseRowModel>> dataMap = new HashMap<>(1);
        dataMap.put(sheetName, data);
        StyleHandler sh = new StyleHandler(dataMap);
        sh.setHandler(handler);
        return writeExcel(fileOutputStream, sheetName, needHead, data, sh);
    }

    private static boolean writeExcel(FileOutputStream fileOutputStream, String sheetName, boolean needHead, List<? extends BaseRowModel> data, WriteHandler handler) throws Exception {
        Map<String, List<? extends BaseRowModel>> dataMap = new HashMap<>(1);
        dataMap.put(sheetName, data);
        return writeExcelWithHandler(fileOutputStream, dataMap, needHead, handler);
    }

    private static boolean writeExcelWithHandler(FileOutputStream fileOutputStream, Map<String, List<? extends BaseRowModel>> data, boolean needHead, WriteHandler handler) throws Exception {
        boolean result = false;
        if (MapUtils.isEmpty(data)) {
            log.warn("write excel data not empty");
            return false;
        }
        com.example.excel.utils.ExcelWriter writer = null;
        try {
            int sheetIndex = 1;
            for (String sheetName : data.keySet()) {
                List<? extends BaseRowModel> models = data.get(sheetName);
                if (models.size() == 0) {
                    continue;
                }
                models = convertData(models);
                writer = com.example.excel.utils.EasyExcelFactory.getWriterWithTempAndHandler(null, fileOutputStream, ExcelTypeEnum.XLSX, needHead, handler);
                Sheet sheet = new Sheet(1, 1, models.get(0).getClass());
                sheetName = StringUtils.isBlank(sheetName) ? "sheet" + sheetIndex : sheetName;
                sheet.setSheetName(sheetName);
                writer.write(models, sheet);
                sheetIndex++;
            }
            result = true;
            log.info("文件生成成功");
        } catch (Exception e) {
            log.error("excel文件生成失败", e);
        } finally {
            if (writer != null) {
                writer.finish();
            }
        }
        return result;
    }

    private static <T> List<T> convertData(List<? extends BaseRowModel> data) throws Exception {
        List<T> result = new ArrayList<>();
        for (Object o : data) {
            final Object copyObj = BeanUtils.transform(o.getClass(), o);

            List<Field> fields = new ArrayList<>();
            Class<?> copyObjClass = copyObj.getClass();
            while (copyObjClass != null) {
                fields.addAll(Arrays.asList(copyObjClass.getDeclaredFields()));
                copyObjClass = copyObjClass.getSuperclass();
            }

            for (Field field : fields) {
                field.setAccessible(true);
                final ExcelValueFormat valueFormat = field.getDeclaredAnnotation(ExcelValueFormat.class);
                if (valueFormat != null) {
                    final String format = valueFormat.format();
                    final Object value = field.get(copyObj);
                    if (value == null) {
                        field.set(copyObj, StringUtils.EMPTY);
                    } else {
                        final String newValue = MessageFormat.format(format, value);
                        field.set(copyObj, newValue);
                    }
                }
            }
            result.add((T) copyObj);
        }
        return result;
    }

    /**
     * 自定义head，sheet 生成excle表
     *
     * @param filePath 绝对路径
     * @param data     数据
     * @param sheet    excle页面样式
     */
    private static void writeWithTemplateAndSheet(String filePath, List<? extends BaseRowModel> data, Sheet sheet) {
        if (CollectionUtils.isEmpty(data)) {
            return;
        }

        sheet = (sheet != null) ? sheet : initSheet;
        sheet.setClazz(data.get(0).getClass());
        OutputStream outputStream = null;
        ExcelWriter writer = null;
        try {
            outputStream = new FileOutputStream(filePath);
            writer = EasyExcelFactory.getWriter(outputStream);
            writer.write(data, sheet);
        } catch (FileNotFoundException e) {
            System.err.println("找不到文件或文件路径错误, 文件：{}" + filePath);
        } finally {
            try {
                if (writer != null) {
                    writer.finish();
                }
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                System.err.println("excel文件导出失败, 失败原因:" + e);
            }
        }
    }
}
