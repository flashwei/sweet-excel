package com.github.excel.model;

import com.github.excel.exception.ExcelReadException;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:03 下午
 * @Description: Excel 模型通用基础类
 */
public class ExcelBaseModel{
	/**
	 * 执行检查方法由子类重写
	 * @throws ExcelReadException
	 */
	public void validator() throws ExcelReadException{

	}
}
