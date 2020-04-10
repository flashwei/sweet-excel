package com.github.excel.model;

import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:05 下午
 * @Description: 导入模板缓存模型
 */
@Data
public class ExcelImportTemplateCacheModel {
	private Integer sheetIndex ;
	private Integer colIndex ;
	private Integer rowIndex ;
	private String text ;
}
