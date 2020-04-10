package com.github.excel.read.handler;

import com.github.excel.read.AbstractExcelReader;
import com.github.excel.model.ExcelReadErrorMsgModel;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:10 下午
 * @Description: 解析handler
 */
public interface ExcelParseHandler {

	ExcelReadErrorMsgModel process(List<AbstractExcelReader.ExcelReadModelDto> readModelDto, boolean endReturnException, boolean readPicture, String template);
}
