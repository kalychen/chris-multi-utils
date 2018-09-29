package com.chris.multi.poi.xls;

import java.lang.annotation.*;

/**
 * Created by Chris Chen
 * 2018/09/18
 * Explain: 数据表头列名
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface XlsColumn {
    String value();//映射列名

    int width() default -1;//列宽 默认-1,框架将按照标题行的字符长度来适配

    boolean required() default false;//是否必填
}
