package com.github.excel.exception;

import java.util.function.Supplier;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:01 下午
 * @Description: 压缩文件异常
 */
public class CompressException extends RuntimeException implements Supplier<CompressException> {
	public CompressException(String message) {
		super(message);
	}

	@Override
	public CompressException get() {
		return this;
	}
}
