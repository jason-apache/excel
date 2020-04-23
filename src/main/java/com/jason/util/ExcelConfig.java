package com.jason.util;

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
     * 导出工作簿名称
     */
    public final static String SHEET_NAME = "Export";
    /**
     * 导出初始化工作簿大小
     */
    public final static int EXPORT_WORK_SIZE = 1000;

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
    }

    /**
     * 布尔型模板
     */
    public final static String IMPORT_TRUE = "是";
    public final static String IMPORT_FALSE = "否";
    public final static String EXPORT_TRUE = "是";
    public final static String EXPORT_FALSE = "否";


    /**
     * 数据模板，仅支持String类型，主要用于字典数据转换、
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
