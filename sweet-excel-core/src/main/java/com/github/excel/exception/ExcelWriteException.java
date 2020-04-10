package com.github.excel.exception;

import java.util.function.Supplier;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:01 下午
 * @Description: 导出excel异常
 */
public class ExcelWriteException extends RuntimeException implements Supplier<ExcelWriteException> {
	public ExcelWriteException(String message) {
		super(message);
	}

	@Override
	public ExcelWriteException get() {
		return this;
	}
}
