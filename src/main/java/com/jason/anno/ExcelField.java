package com.jason.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * TODO
 *
 * @Author jason
 * @createTime 2019年12月18日 21:12
 * @Description
 * 优先级：position > title > 默认（不使用该注解时，默认以字段名映射excel）
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD,ElementType.TYPE,ElementType.FIELD})
public @interface ExcelField {

    //对应excel列名
    String title() default "";

    //从第几行开始读取
    int startRow() default -1;

    //从第几个工作簿开始读取
    int startSheet() default -1;

    String sheetName() default "";

    //指定excel具体某一列给实体类字段赋值
    int position() default -1;

    //当使用引用类型时目标方法的名称
    String targetMethod() default "";

    //当使用引用类型时目标方法的参数
    Class<?> targetClass() default String.class;

    //导出排序
    int sort() default -1;

    //导出title
    String excelTile() default "";

    //是否使用模板
    boolean useTemplate() default false;

    //使用模板中的第几个map
    int templatePosition() default 0;

    //是否导入
    boolean isImport() default true;
}
