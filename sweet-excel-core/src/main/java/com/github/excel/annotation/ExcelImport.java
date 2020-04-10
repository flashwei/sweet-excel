package com.github.excel.annotation;

import java.lang.annotation.*;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 6:56 下午
 * @Description: Excel 读取
 */
@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelImport {
	/**
	 * 启用分隔符
	 */
	boolean enableSeparator() default true ;
}
