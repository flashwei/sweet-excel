package com.github.model;

import com.github.excel.annotation.ExcelImport;
import com.github.excel.annotation.ExcelImportProperty;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ReadPictureModel;
import lombok.Data;

import java.util.Date;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
@Data
@ExcelImport(enableSeparator = false)
public class UserExcelDtoImportList1 extends ExcelBaseModel {
	@ExcelImportProperty(titleName = "姓名")
	private String name;
	@ExcelImportProperty(titleName = "性别")
	private String sex;
	@ExcelImportProperty(titleName = "年龄")
	private Integer age;
	@ExcelImportProperty(titleName = "身高")
	private Float height;
	@ExcelImportProperty(titleName = "昵称")
	private String nickName;
	@ExcelImportProperty(titleName = "头像")
	private String avater;
	@ExcelImportProperty(titleName = "创建时间",formatPattern = "yyyy-MM-dd")
	private Date createTime;
	@ExcelImportProperty(titleName = "备注")
	private String remark;
	@ExcelImportProperty(titleName = "国籍")
	private String contry;
	@ExcelImportProperty(titleName = "logo")
	private List<ReadPictureModel> logo;

	@Override
	public String toString() {
		return "UserExcelDtoImportList1{" + "name='" + name + '\'' + ", sex='" + sex + '\'' + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", createTime=" + createTime + ", remark='" + remark + '\'' + ", contry='" + contry + '\'' + '}';
	}
}
