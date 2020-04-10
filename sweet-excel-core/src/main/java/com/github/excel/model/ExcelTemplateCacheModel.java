package com.github.excel.model;

import com.github.excel.boot.WorkbookCachePool;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:08 下午
 * @Description: Excel 文件缓存model
 */
@AllArgsConstructor
public class ExcelTemplateCacheModel {

	@Getter
	private final ThreadLocal<WorkbookCachePool.WorkbookCacheModel> workbookThreadLocal;

}
