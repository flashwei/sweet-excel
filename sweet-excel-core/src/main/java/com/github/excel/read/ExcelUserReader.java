package com.github.excel.read;

import com.github.excel.read.handler.ExcelParseHandler;
import com.github.excel.read.handler.impl.ExcelUserParseHandler;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:14 下午
 * @Description: excel 事件模式下读取
 */
@Slf4j
public class ExcelUserReader extends AbstractExcelReader{

	public ExcelUserReader(String excelName, InputStream readExcelStream) {
		super(excelName, readExcelStream);
	}
	public ExcelUserReader(String excelName, InputStream readExcelStream,String template) {
		super(excelName, readExcelStream,template);
	}

	@Override
	public ExcelParseHandler createHandler() {
		return new ExcelUserParseHandler(readExcelStream,excelName);
	}
}
