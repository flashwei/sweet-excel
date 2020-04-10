package com.github.excel.read.handler.impl;

import com.github.excel.read.ExcelReaderDataFormat;
import com.github.excel.read.handler.ExcelParseHandler;
import com.google.common.base.Throwables;
import com.google.common.collect.Maps;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.model.ExcelReadErrorMsgModel;
import com.github.excel.read.ExcelDefaultReaderDataFormat;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:10 下午
 * @Description: excel 读取抽象类
 */
@Slf4j
public abstract class AbstractExcelParseHandler  implements ExcelParseHandler {
	/**
	 * 读取excel文件
	 */
	protected InputStream readExcelStream;
	protected String excelName;
	protected ExcelReaderDataFormat defaultFormat = new ExcelDefaultReaderDataFormat();
	protected ExcelReadErrorMsgModel errorMsgModel = new ExcelReadErrorMsgModel();
	protected Map<Class<? extends ExcelReaderDataFormat>, ExcelReaderDataFormat> dataFormatMap = Maps.newHashMap();

	public AbstractExcelParseHandler(InputStream readExcelStream, String excelName) {
		this.readExcelStream = readExcelStream;
		this.excelName = excelName;
	}

	public ExcelReaderDataFormat getDataFormatThenCache(Class<? extends ExcelReaderDataFormat> formatCla) {
		if (formatCla == ExcelDefaultReaderDataFormat.class) {
			return defaultFormat;
		} else {
			ExcelReaderDataFormat excelReaderDataFormat = dataFormatMap.get(formatCla);
			if (excelReaderDataFormat != null) {
				return excelReaderDataFormat;
			}
			try {
				excelReaderDataFormat = formatCla.newInstance();
			} catch (InstantiationException e) {
				log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
				throw new ExcelReadException(e.getMessage());
			} catch (IllegalAccessException e) {
				log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
				throw new ExcelReadException(e.getMessage());
			}
			dataFormatMap.put(formatCla, excelReaderDataFormat);
			return excelReaderDataFormat;
		}

	}
}
