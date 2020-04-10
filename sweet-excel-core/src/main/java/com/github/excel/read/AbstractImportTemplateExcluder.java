package com.github.excel.read;

import com.github.excel.util.StringUtil;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelReadException;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:12 下午
 * @Description: 导入模板排除器
 */
public abstract class AbstractImportTemplateExcluder {

	private Map<String,Map<String,String>> templateExcludeMap = new HashMap<>();

	public <T extends ImportExclude> boolean isExclude(T exclude) {
		return exclude.isExclude();
	}

	public AbstractImportTemplateExcluder addExclude(ImportExclude importExclude) {
		if (Objects.isNull(importExclude)) {
			return this;
		}
		importExclude.addExclude();
		return this;
	}

	private Map<String, String> getExcludeMapByTemplate(String template) {
		Map<String, String> excludeMap = templateExcludeMap.get(template);
		if (Objects.isNull(excludeMap)) {
			excludeMap = new HashMap<>();
			templateExcludeMap.put(template, excludeMap);
		}
		return excludeMap;
	}

	private boolean hasExclude(String template,String key) {
		Map<String, String> excludeMap = templateExcludeMap.get(template);
		if (Objects.isNull(excludeMap)) {
			return false ;
		}
		if (Objects.nonNull(excludeMap.get(key))) {
			return true;
		}
		return false ;
	}

	public abstract void addTemplateExclude();

	public interface ImportExclude {
		// 添加排除
		void addExclude();
		// 是否排除
		boolean isExclude();
	}

	@Data
	@AllArgsConstructor
	public class ColumnExclude implements ImportExclude {
		private Integer colIndex ;
		private Integer rowIndex ;
		private String sheetName ;
		private String templateName;

		@Override
		public void addExclude() {
			if (StringUtil.isEmpty(this.getSheetName()) || Objects.isNull(this.getRowIndex()) || Objects.isNull(this.getColIndex()) || StringUtil.isEmpty(templateName)) {
				throw new ExcelReadException("Column exclude params can't be null");
			}
			Map<String, String> excludeMap = getExcludeMapByTemplate(templateName);
			excludeMap.put(this.getSheetName() + ExcelConstant.SHORT_TERM + this.getRowIndex() + ExcelConstant.SHORT_TERM + this.getColIndex(), ExcelConstant.NULL_STR);
		}

		@Override
		public boolean isExclude(){
			String key = this.getSheetName() + ExcelConstant.SHORT_TERM + this.getRowIndex() + ExcelConstant.SHORT_TERM + this.getColIndex();
			return hasExclude(templateName,key);
		}
	}

	@Data
	@AllArgsConstructor
	public class SheetExclude implements ImportExclude {
		private String sheetName ;
		private String templateName;

		@Override
		public void addExclude() {
			if (Objects.isNull(this.getSheetName())|| StringUtil.isEmpty(templateName)) {
				throw new ExcelReadException("Column exclude params can't be null");
			}
			Map<String, String> excludeMap = getExcludeMapByTemplate(templateName);
			excludeMap.put(this.getSheetName(),ExcelConstant.NULL_STR);
		}

		@Override
		public boolean isExclude() {
			return hasExclude(templateName,this.getSheetName());
		}
	}

	@Data
	@AllArgsConstructor
	public class RowExclude implements ImportExclude {
		private Integer index ;
		private String sheetName ;
		private String templateName;

		@Override
		public void addExclude() {
			if (Objects.isNull(this.getSheetName()) || Objects.isNull(this.getIndex()) || StringUtil.isEmpty(templateName)) {
				throw new ExcelReadException("Column exclude params can't be null");
			}
			Map<String, String> excludeMap = getExcludeMapByTemplate(templateName);
			excludeMap.put(this.getSheetName() + ExcelConstant.SHORT_TERM + this.getIndex(),ExcelConstant.NULL_STR);
		}

		@Override
		public boolean isExclude() {
			String key = this.getSheetName() + ExcelConstant.SHORT_TERM + this.getIndex();
			return hasExclude(templateName,key);
		}
	}
}
