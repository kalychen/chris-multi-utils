package com.chris.multi.poi.xls;

import java.lang.annotation.*;

/**
 * Created by Chris Chen
 * 2018/09/18
 * Explain: 数据表名
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface XlsSheet {
    String value();

    int maxLines() default 65534;//单表支持的最大数据行数

    String ext() default "-000";//分表后缀格式
}
