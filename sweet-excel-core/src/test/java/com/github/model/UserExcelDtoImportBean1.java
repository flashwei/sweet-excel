package com.github.model;

import com.github.excel.annotation.ExcelImport;
import com.github.excel.annotation.ExcelImportProperty;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 简单导出
 */
@Data
@ExcelImport
public class UserExcelDtoImportBean1 extends ExcelBaseModel {
	@ExcelImportProperty(titleName = "姓名")
	private String name;
	@ExcelImportProperty(titleName = "性别",checkNull = true)
	private Byte sex;
	@ExcelImportProperty(titleName = "年龄")
	private Integer age;
	@ExcelImportProperty(titleName = "身高")
	private Float height;
	@ExcelImportProperty(titleName = "昵称",checkNull = true)
	private String nickName;
	@ExcelImportProperty(titleName = "头像")
	private String avater;
	/*@ExcelImportProperty(titleName = "logo")
	private byte[] logo;*/
	@ExcelImportProperty(titleName = "创建时间",formatPattern = "yyyy-MM-dd")
	private Date createTime;
	@ExcelImportProperty(titleName = "备注",checkNull = true)
	private String remark;
	@ExcelImportProperty(titleName = "国籍",checkNull = true)
	private String contry;

	@Override
	public String toString() {
		return "UserExcelDtoImportBean1{" + "name='" + name + '\'' + ", sex=" + sex + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", createTime=" + createTime + ", remark='" + remark + '\'' + ", contry='" + contry + '\'' + '}';
	}
}
