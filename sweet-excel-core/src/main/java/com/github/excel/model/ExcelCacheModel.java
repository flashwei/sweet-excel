package com.github.excel.model;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import lombok.Data;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:03 下午
 * @Description: 导出缓存model
 */
@Data
public class ExcelCacheModel {
	/**
	 * excelExport 注解
	 */
	private ExcelExport excelExport;
	/**
	 * 字段列表
	 */
	private List<ExcelCacheFieldModel> fieldModelList;
	/**
	 * 字段map
	 */
	private Map<String,ExcelCacheFieldModel> fieldModelMap;

	@Data
	public static class ExcelCacheFieldModel {
		/**
		 * exportCell 注解
		 */
		private ExcelExportCell exportCell;
		/**
		 * get 函数
		 */
		private Method getMethod ;
		/**
		 * 字段名称
		 */
		private String fieldName ;
		/**
		 * 标题样式名称
		 */
		private String titleStyleName ;
		/**
		 * 内容样式名称
		 */
		private String contentStyleName ;
		/**
		 * 偶数行样式名称
		 */
		private String evenRowStyleName ;
		/**
		 * 标题默认高度
		 */
		private Short titleRowHeight;
		/**
		 * 内容默认高度
		 */
		private Short contentRowHeight;
		/**
		 * 单元格默认宽度
		 */
		private Short colWidth;
		/**
		 * 是否map类型
		 */
		private boolean isMap;
	}
}
