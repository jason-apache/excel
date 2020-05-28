package com.jason.util;

import org.apache.poi.ss.usermodel.*;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Jason
 * @Date: 2020/3/30 13:32
 * @Description:
 */
public class ExcelConfig {

    /**
     * 日期格式
     */
    public final static String DATE_IMPORT_FORMAT = "yyyy-MM-dd";
    public final static String DATE_EXPORT_FORMAT = "yyyy-MM-dd";

    /**
     * call分隔符
     */
    public final static String CALL_SEPARATOR = ".";

    /**
     * 导入默认集合大小
     */
    public final static Integer DEFAULT_COLLECTION_SIZE = 1000;

    /**
     * 导出工作簿名称
     */
    public final static String SHEET_NAME = "Export";
    /**
     * 导出初始化工作簿大小
     */
    public final static int EXPORT_WORK_SIZE = 1000;
    /**
     *
     */
    public final static Integer EXPORT_FORMAT_RANGE = Short.MAX_VALUE * 2 * 10;

    /**
     * 样式常量
     */
    public static class Style {
        /**
         * 标题样式
         */
        public final static String TITLE = "headTitle";
        /**
         * 首行样式
         */
        public final static String HEAD_ROW = "headRow";
        /**
         * 单元格样式
         */
        public final static String DEFAULT_STYLE = "data";
        /**
         * 字体类型
         */
        public final static String FONT_NAME = "微软雅黑";
        /**
         * 标题字体大小
         */
        public final static short FONT_TITLE_SIZE = 18;
        /**
         * headRow字体大小
         */
        public final static short FONT_HEAD_SIZE = 15;
        /**
         * 单元格字体大小
         */
        public final static short FONT_CELL_SIZE = 12;
        /**
         * 单元格默认最小宽度
         */
        public final static int CELL_MIN_WIDTH = 4000;

        public static Map<String,CellStyle> getDataStyle(Workbook workbook){
            Map<String, CellStyle> styles = new HashMap<>(1);
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.LEFT);
            style.setBorderRight(BorderStyle.THIN);
            style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderLeft(BorderStyle.THIN);
            style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderTop(BorderStyle.THIN);
            style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            Font dataFont = workbook.createFont();
            dataFont.setFontName(ExcelConfig.Style.FONT_NAME);
            dataFont.setFontHeightInPoints(ExcelConfig.Style.FONT_CELL_SIZE);
            style.setFont(dataFont);
            styles.put(ExcelConfig.Style.DEFAULT_STYLE, style);
            return styles;
        }

        public static Map<String,CellStyle> getHeadRowStyle(Workbook workbook) {
            Map<String, CellStyle> styles = new HashMap<>(1);
            CellStyle style = workbook.createCellStyle();
            //设置headRow样式
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setBorderRight(BorderStyle.THIN);
            style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderLeft(BorderStyle.THIN);
            style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderTop(BorderStyle.THIN);
            style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            Font headRowFont = workbook.createFont();
            headRowFont.setFontName(ExcelConfig.Style.FONT_NAME);
            headRowFont.setFontHeightInPoints(ExcelConfig.Style.FONT_HEAD_SIZE);
            headRowFont.setBold(true);
            style.setFont(headRowFont);
            styles.put(ExcelConfig.Style.HEAD_ROW, style);
            return styles;
        }

        public static Map<String,CellStyle> getTitleStyle(Workbook workbook){
            Map<String, CellStyle> styles = new HashMap<>(1);
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setBorderRight(BorderStyle.THIN);
            style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderLeft(BorderStyle.THIN);
            style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderTop(BorderStyle.THIN);
            style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            Font titleFont = workbook.createFont();
            titleFont.setFontName(ExcelConfig.Style.FONT_NAME);
            titleFont.setFontHeightInPoints(ExcelConfig.Style.FONT_TITLE_SIZE);
            titleFont.setBold(true);
            style.setFont(titleFont);
            styles.put(ExcelConfig.Style.TITLE, style);
            return styles;
        }

        public static Map<String, CellStyle> getStyles(Workbook workbook){
            Map<String, CellStyle> styles = new HashMap<>(3);
            styles.put(Style.DEFAULT_STYLE,getDataStyle(workbook).get(Style.DEFAULT_STYLE));
            styles.put(Style.HEAD_ROW,getHeadRowStyle(workbook).get(Style.HEAD_ROW));
            styles.put(Style.TITLE,getTitleStyle(workbook).get(Style.TITLE));
            return styles;
        }
    }

    /**
     * 布尔型模板
     */
    public final static String IMPORT_TRUE = "是";
    public final static String IMPORT_FALSE = "否";
    public final static String EXPORT_TRUE = "是";
    public final static String EXPORT_FALSE = "否";

    /**
     * 模板key常量
     */
    public final static String DEFAULT_NAME_KEY = "default";
    /**
     * 配置字典数据转换，仅支持String类型
     */
    public enum Template{

        TEST1("0","测试单元1"),
        TEST2("1","测试单元2"),
        TEST3("2","测试单元3"),
        TEST4("3","测试单元4"),
        LAST("999","测试单元1000"),
        ;

        private String code;
        private String title;

        Template(String code, String title) {
            this.code = code;
            this.title = title;
        }

        public String getCode() {
            return code;
        }

        public Template setCode(String code) {
            this.code = code;
            return this;
        }

        public String getTitle() {
            return title;
        }

        public Template setTitle(String title) {
            this.title = title;
            return this;
        }
    }

    public static Map<String,String> getTemplateTitle(){
        Template[] values = Template.values();
        Map<String, String> map = new HashMap<>(values.length);
        for (Template value : values) {
            map.put(value.getCode(), value.getTitle());
        }
        return map;
    }

    public static Map<String,String> getTemplateCode(){
        Template[] values = Template.values();
        Map<String, String> map = new HashMap<>(values.length);
        for (Template value : values) {
            map.put(value.getTitle(), value.getCode());
        }
        return map;
    }

    //默认值
    public static class DefaultValue{
        public static String defaultString = null;
        public static Integer defaultInteger = -1;
        public static Long defaultLong = -1L;
        public static Short defaultShort = -1;
        public static Byte defaultByte = -1;
        public static Character defaultChar = null;
        public static Float defaultFloat = -1.1F;
        public static Double defaultDouble = -1.1D;
        public static Boolean defaultBoolean = null;
        public static Date defaultDate = null;
    }
}
