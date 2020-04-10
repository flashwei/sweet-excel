package com.github.excel.write;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:17 下午
 * @Description: 自定义写接口
 */
@FunctionalInterface
public interface ExcelCustomWriter {
	void execute(Workbook workbook);
}
