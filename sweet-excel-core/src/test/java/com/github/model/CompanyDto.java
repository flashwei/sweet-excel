package com.github.model;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

import java.util.Calendar;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:28 下午
 * @Description: 测试demo
 */
@Data
@ExcelExport
public class CompanyDto extends ExcelBaseModel {
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
	@ExcelExportCell(titleName = "创建人")
	private UserExcelDto createUser ;

}
