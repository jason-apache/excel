package com.jason.util;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @Auther: Jason
 * @Date: 2020/3/30 13:32
 * @Description:
 */
public class ExcelConfig {

    //日期格式
    public final static String DATE_IMPORT_FORMAT = "yyyy-MM-dd";
    public final static String DATE_EXPORT_FORMAT = "yyyy-MM-dd";

    //工作簿名称
    public final static String SHEET_NAME = "Export";

    //样式key常量
    public static class Style {
        public final static String HEAD_ROW = "headRow";
        public final static String HEAD_TITLE = "headTitle";
        public final static String DEFAULT_STYLE = "default";
        public final static String FONT_NAME = "微软雅黑";
    }

    //布尔型模板
    public final static String IMPORT_TRUE = "是";
    public final static String IMPORT_FALSE = "否";
    public final static String EXPORT_TRUE = "是";
    public final static String EXPORT_FALSE = "否";

    //其他类型模板
    //仅支持String类型，主要用于字典数据转换
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
        Map<String, String> map = new HashMap<>();
        Template[] values = Template.values();
        for(int i = 0 ; i < values.length; i++){
            map.put(values[i].getCode(),values[i].getTitle());
        }
        return map;
    }

    public static Map<String,String> getTemplateCode(){
        Map<String, String> map = new HashMap<>();
        Template[] values = Template.values();
        for(int i = 0 ; i < values.length; i++){
            map.put(values[i].getTitle(),values[i].getCode());
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
