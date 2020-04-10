package com.github.excel.model;

import com.github.excel.annotation.ExcelImport;
import com.github.excel.annotation.ExcelImportProperty;
import lombok.Data;

import java.lang.reflect.Method;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:03 下午
 * @Description: 导入缓存model
 */
@Data
public class ExcelCacheImportModel {
	private ExcelImport excelImport;
	private Map<String, ExcelCacheImportFieldModel> fieldModelMap;

	@Data
	public static class ExcelCacheImportFieldModel {
		private ExcelImportProperty importProperty;
		private Method setMethod;
	}
}
