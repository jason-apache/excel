package com.jason.anno;

import com.jason.util.ExcelConfig;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 *
 * @Author jason
 * @createTime 2019年12月18日 21:12
 * @Description
 * 优先级：position > title > 默认（不使用该注解时，默认以字段名映射excel）
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD,ElementType.TYPE,ElementType.FIELD})
public @interface ExcelField {
    /**
     * 对应excel列名
     */
    String title() default "";
    /**
     *从第几行开始读取
     */
    int startRow() default -1;
    /**
     *从第几个工作簿读取
     */
    int sheetIndex() default -1;
    /**
     *工作簿名
     */
    String sheetName() default "";
    /**
     *指定excel具体某一列给实体类字段赋值
     */
    int position() default -1;
    /**
     * 调用属性，该操作递归反射对象，比较占用系统资源   eg: parent.name
     * 可以配合callMethod使用  eg: call = "parent.student",callMethod = "setName", callClass = String.class
     */
    String call() default "";
    /**
     *调用方法名
     */
    String callMethod() default "";
    /**
     *调用方法参数类型
     */
    Class<?> callClass() default String.class;
    /**
     *导出排序
     */
    int sort() default -1;
    /**
     *是否使用模板
     */
    boolean useTemplate() default false;
    /**
     *使用模板的key
     */
    String templateNameKey() default ExcelConfig.DEFAULT_NAME_KEY;
    /**
     *是否导入
     */
    boolean isImport() default true;
}
