package com.github.excel.exception;

import java.util.function.Supplier;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:01 下午
 * @Description: 读取excel异常
 */
public class ExcelReadException extends RuntimeException implements Supplier<ExcelReadException> {
	public ExcelReadException(String message) {
		super(message);
	}

	@Override
	public ExcelReadException get() {
		return this;
	}
}
