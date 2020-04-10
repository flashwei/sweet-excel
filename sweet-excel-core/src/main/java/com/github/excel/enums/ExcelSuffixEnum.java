package com.github.excel.enums;

import com.github.excel.constant.ExcelConstant;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:00 下午
 * @Description: Excel后缀
 */
public enum ExcelSuffixEnum {

	XLS(ExcelConstant.XLS_STR),

	XLSX(ExcelConstant.XLSX_STR);

	ExcelSuffixEnum(String suffix){
		this.suffix = suffix;
	}

	private String suffix ;

	public String getSuffix() {
		return suffix;
	}
}
