package com.github.excel.read;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.handler.impl.ExcelEventXlsParseHandler;
import com.github.excel.read.handler.impl.ExcelEventXlsxParseHandler;

import java.io.InputStream;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:14 下午
 * @Description: Excel 读取器创建工厂
 */
public class ExcelReaderFactory {

	/**
	 * 创建Excel读取
	 *
	 * @param excelName
	 * @param readExcelStream
	 * @return
	 */
	public static ExcelReader createUserReader(String excelName, InputStream readExcelStream) {
		return new ExcelUserReader(excelName, readExcelStream);
	}

	/**
	 * 创建Excel读取器
	 * @param excelName 文件名称
	 * @param readExcelStream 流
	 * @param template 模板名
	 * @return
	 */
	public static ExcelReader createUserReader(String excelName, InputStream readExcelStream,String template) {
		return new ExcelUserReader(excelName, readExcelStream,template);
	}

	/**
	 *  创建Excel事件模式下读取
	 * @param fileName
	 * @param <T>
	 * @return
	 */
	public static <T> ExcelEventReader<T> createExcelEventReader(String fileName) {
		ExcelEventReader excelReader;
		// 处理excel03文件
		if (fileName.endsWith(ExcelConstant.XLS_STR)) {
			excelReader = new ExcelEventXlsParseHandler();
			// 处理excel07文件
		} else if (fileName.endsWith(ExcelConstant.XLSX_STR)) {
			excelReader = new ExcelEventXlsxParseHandler();
		} else {
			throw new ExcelReadException("file.extension.invalid");
		}
		return excelReader;
	}
}
