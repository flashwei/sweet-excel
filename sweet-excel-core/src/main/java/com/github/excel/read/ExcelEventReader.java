package com.github.excel.read;

import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.handler.AbstractExecuteHandler;

import java.io.InputStream;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:12 下午
 * @Description: excel 事件模式下读取
 */
public interface ExcelEventReader<T> {

	/**
	 * 设置行解析器
	 * @param rowReader
	 */
	void setRowReader(ExcelEventRowReader<T> rowReader);

	/**
	 * 设置执行处理器
	 * @param executeHandler
	 */
	void setExecuteHandler(AbstractExecuteHandler<T> executeHandler);

	/**
	 * 解析操作
	 * @param inputStream
	 */
	void process(InputStream inputStream) throws ExcelReadException;

}
