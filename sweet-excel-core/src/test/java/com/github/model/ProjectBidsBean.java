package com.github.model;

import com.github.excel.annotation.ExcelImport;
import com.github.excel.annotation.ExcelImportProperty;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;
import lombok.ToString;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 简单导出
 */
@Data
@ExcelImport(enableSeparator = false)
@ToString
public class ProjectBidsBean extends ExcelBaseModel {
	@ExcelImportProperty(titleName = "规划标题")
	private String title;
	@ExcelImportProperty(titleName = "招标类别")
	private String type;
	@ExcelImportProperty(titleName = "招标模式")
	private String model;
	@ExcelImportProperty(titleName = "招标方式")
	private String bidsMethod;
	@ExcelImportProperty(titleName = "包干方式")
	private String doneMethod;
	@ExcelImportProperty(titleName = "标段数量")
	private Integer num;
	@ExcelImportProperty(titleName = "工期(天)")
	private Integer workTotalDays;
}
