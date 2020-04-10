package com.github.excel.util;

import com.github.excel.constant.ExcelConstant;

/**
 * @description: 字符串工具类
 * @author: Vachel Wang
 * @create: 2019-09-07 16:28
 **/
public class StringUtil {
	public static boolean isEmpty(Object str) {
		return (str == null || ExcelConstant.NULL_STR.equals(str));
	}
	public static boolean notEmpty(Object str) {
		return !isEmpty(str);
	}
}
