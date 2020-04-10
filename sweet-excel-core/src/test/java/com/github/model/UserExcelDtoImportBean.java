package com.github.model;

import com.github.excel.annotation.ExcelImport;
import com.github.excel.annotation.ExcelImportProperty;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ReadPictureModel;
import lombok.Data;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 简单导出
 */
@Data
@ExcelImport
public class UserExcelDtoImportBean extends ExcelBaseModel {
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
	@ExcelImportProperty(titleName = "logo",checkNull = true)
	private ReadPictureModel logo;
	@ExcelImportProperty(titleName = "创建时间",formatPattern = "yyyy-MM-dd")
	private Date createTime;
	@ExcelImportProperty(titleName = "备注",checkNull = true)
	private String remark;
	@ExcelImportProperty(titleName = "余额",checkNull = true)
	private BigDecimal money;
	@ExcelImportProperty(titleName = "余额BigInteger",checkNull = true)
	private BigInteger moneyInt;
	@ExcelImportProperty(titleName = "小孩",checkNull = true)
	private Boolean hasChlid;

	@Override
	public String toString() {
		return "UserExcelDtoImportBean{" + "name='" + name + '\'' + ", sex=" + sex + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", logo=" + logo + ", createTime=" + createTime + ", remark='" + remark + '\'' + ", money=" + money + ", moneyInt=" + moneyInt + '}';
	}

	@Override
	public void validator() throws ExcelReadException {
		if (sex == null) {
			throw new ExcelReadException("123");
		}
	}
}
