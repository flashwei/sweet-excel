package com.github.model;

import com.github.export.ExcelDefaultDataFormatTest;
import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelExportCellTitleModelEnum;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.enums.ExcelExportListFillTypeEnum;
import com.github.excel.enums.ExcelExportScopeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelBasicStyle;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 测试demo
 */
@Data
@ExcelExport(rowIndex = 13, colIndex = 0,fillType = ExcelExportListFillTypeEnum.SHIFT,scope = ExcelExportScopeEnum.CURRENT_SHEET,fillStyle = ExcelExportFillStyleEnum.HORIZONTAL, nameSpace = "user",titleStyleName = ExcelBasicStyle.STYLE_LIST_TITLE,contentStyleName = ExcelBasicStyle.STYLE_CONTENT/*,mergeTitleRowNum = 1,mergeTitleColNum = 1,mergeContentColNum = 1,mergeContentRowNum = 1*/)
public class UserExcelDto extends ExcelBaseModel {
	@ExcelExportCell(titleName = "姓名", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE)
	private String name;
	@ExcelExportCell(titleName = "性别", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE)
	private Byte sex;
	@ExcelExportCell(titleName = "年龄", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE)
	private Integer age;
	@ExcelExportCell(titleName = "年龄short", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE)
	private Short ageShort;
	@ExcelExportCell(titleName = "身高double",colWidth=-1/*, titleStyleName = ExcelBasicStyle.STYLE_TITLE, contentStyleName = ExcelBasicStyle.STYLE_CONTENT*/, titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private Double heightDouble;
	@ExcelExportCell(titleName = "身高", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE)
	private Float height;
	@ExcelExportCell(titleName = "昵称",colWidth = -1, titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE,fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private String nickName;
	@ExcelExportCell(titleName = "头像", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, linkName = "点击查看", colWidth = -1)
	private String avater;
	@ExcelExportCell(titleName = "邮箱", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE,  linkName = "点击发送", colWidth = -1)
	private String email;
	@ExcelExportCell(titleName = "账户金额", colWidth = 200/*, titleStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT, contentStyleName = ExcelBasicStyle.STYLE_CONTENT*/, titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, separator = " | ")
	private Long money;
	@ExcelExportCell(titleName = "金额格式化", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, formatPattern = "###,###,###.000")
	private BigDecimal moneyBig;
	@ExcelExportCell(titleName = "锁定", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE)
	private Boolean lock;
	@ExcelExportCell(titleName = "创建时间", formatPattern = "yyyy-MM-dd HH:mm:ss.sss", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, formatter = ExcelDefaultDataFormatTest.class, colWidth = -1,fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private Date createTime;
	@ExcelExportCell(titleName = "修改时间", formatPattern = "yyyy-MM-dd", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE,colWidth = 300,fillStyle = ExcelExportFillStyleEnum.VERTICAL,verticalNewLine = true)
	private Calendar updateTime;
	@ExcelExportCell(titleName = "logo", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE,/* contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ,*/ colIndex = 0, rowIndex = 0, rowHeight = 51, colWidth = 430)
	private byte[] logo;
	@ExcelExportCell(titleName ="公司")
	private CompanyDto company;
}
