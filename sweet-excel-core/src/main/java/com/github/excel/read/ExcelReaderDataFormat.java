package com.github.excel.read;

import java.text.ParseException;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:13 下午
 * @Description: excel读取格式化
 */
@FunctionalInterface
public interface ExcelReaderDataFormat {
	Object format(Object data, String pattern, Class<?> targetCla) throws ParseException;
}
