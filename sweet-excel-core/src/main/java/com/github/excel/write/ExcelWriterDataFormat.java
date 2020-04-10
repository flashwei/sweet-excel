package com.github.excel.write;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:18 下午
 * @Description: 导出格式化
 */
@FunctionalInterface
public interface ExcelWriterDataFormat {
	Object format(Object data, String pattern);
}
