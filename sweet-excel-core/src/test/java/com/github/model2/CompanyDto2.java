package com.github.model2;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.export.ExcelCustomStyle;
import lombok.Data;

import java.util.Calendar;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
@Data
@ExcelExport(theme = ExcelThemeEnum.ZEBRA,incrementSequenceNo = false,incrementSequenceTitle = "公司编号")
public class CompanyDto2 extends ExcelBaseModel {
	@ExcelExportCell(titleName = "公司名称")
	private String name;
	@ExcelExportCell(titleName = "公司地址")
	private String address;
	@ExcelExportCell(titleName = "公司创建时间")
	private Date createTime;
	@ExcelExportCell(titleName = "修改时间")
	private Calendar updateTime;
	@ExcelExportCell(titleName = "公司人数")
	private Integer persons;
	@ExcelExportCell(titleName = "LOGO")
	private byte[] logo;

}
