package com.github.model;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ComboBoxModel;
import com.github.excel.model.CommentModel;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.NumberScopeModel;
import com.github.excel.write.ExcelBasicStyle;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 简单导出
 */
@Data
@ExcelExport(theme = ExcelThemeEnum.ZEBRA)
public class UserExcelDto3 extends ExcelBaseModel {
	@ExcelExportCell(titleName = "姓名",fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private String name;
	@ExcelExportCell(titleName = "性别",commentText = "性别只能是男女",commentFontName = ExcelBasicStyle.FONT_SIZE16_BLOLD_RED,disable = false,fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private Byte sex;
	@ExcelExportCell(titleName = "性别-1",fillStyle = ExcelExportFillStyleEnum.VERTICAL,contentStyleName = ExcelBasicStyle.STYLE_FOREGROUND_COLOR_YELLOW)
	private CommentModel sexStr;
	@ExcelExportCell(titleName = "年龄",fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private Integer age;
	@ExcelExportCell(titleName = "身高",disable = true)
	private Float height;
	@ExcelExportCell(titleName = "昵称")
	private String nickName;
	@ExcelExportCell(titleName = "头像")
	private String avater;
	@ExcelExportCell(titleName = "logo")
	private byte[] logo;
	@ExcelExportCell(titleName = "创建时间")
	private Date createTime;
	@ExcelExportCell(titleName = "公司性质",dropDownOptions = {"国企","民营企业"})
	private String companyType;
	@ExcelExportCell(titleName = "公司性质combo",colWidth = 100)
	private ComboBoxModel companyType1;
	@ExcelExportCell(titleName = "年龄范围")
	private NumberScopeModel scopeModel;
}
