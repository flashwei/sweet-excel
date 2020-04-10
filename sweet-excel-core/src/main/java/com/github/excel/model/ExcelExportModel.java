package com.github.excel.model;

import lombok.Data;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:05 下午
 * @Description: 导出model
 */
@Data
public class ExcelExportModel{
	private String sheetName;
	private ExcelBaseModel dataModel;
	private List<? extends ExcelBaseModel> dataModelList;
	private ExcelCacheModel cacheModel;
	private Class<? extends ExcelBaseModel> excelModelClass;
	private Integer rowIndex;
	private Integer colIndex;
	private Boolean fillTemplate;
}
