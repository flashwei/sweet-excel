package com.github.model;

import com.github.export.ExcelDefaultDataFormatTest;
import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelExportCellTitleModelEnum;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.enums.ExcelExportScopeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelBasicStyle;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 测试demo
 */
@Data
@ExcelExport(rowIndex = 1, colIndex = 0, scope = ExcelExportScopeEnum.CURRENT_SHEET, nameSpace = "user")
public class UserExcelDto2 extends ExcelBaseModel {
	@ExcelExportCell(titleName = "姓名", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL, mergeTitleColNum = 1,mergeRowNum = 1,mergeContentColNum = 1)
	private String name;
	@ExcelExportCell(titleName = "性别", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL,mergeTitleColNum = 2,mergeRowNum = 2,mergeContentColNum = 1)
	private Byte sex;
	@ExcelExportCell(titleName = "年龄", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL,mergeTitleColNum = 3,mergeRowNum = 1,mergeContentColNum = 1)
	private Integer age;
	@ExcelExportCell(titleName = "年龄short", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL)
	private Short ageShort;
	@ExcelExportCell(titleName = "身高double", titleStyleName = ExcelBasicStyle.STYLE_TITLE, contentStyleName = ExcelBasicStyle.STYLE_CONTENT, titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL)
	private Double heightDouble;
	@ExcelExportCell(titleName = "身高", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private Float height;
	@ExcelExportCell(titleName = "昵称", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL,mergeTitleColNum = 3,mergeRowNum = 1,mergeContentColNum = 2)
	private String nickName;
	@ExcelExportCell(titleName = "头像", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL, linkName = "点击查看", colWidth = -1)
	private String avater;
	@ExcelExportCell(titleName = "邮箱", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL, linkName = "点击发送", colWidth = -1)
	private String email;
	@ExcelExportCell(titleName = "账户金额", colWidth = 200, titleStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT, contentStyleName = ExcelBasicStyle.STYLE_CONTENT, titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, separator = " | ",fillStyle = ExcelExportFillStyleEnum.HORIZONTAL, colIndex = 20, rowIndex = 1)
	private Long money;
	@ExcelExportCell(titleName = "", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, formatPattern = "###,###,###.000", fillStyle = ExcelExportFillStyleEnum.VERTICAL)
	private BigDecimal moneyBig;
	@ExcelExportCell(titleName = "锁定", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL)
	private Boolean lock;
	@ExcelExportCell(titleName = "创建时间", formatPattern = "yyyy-MM-dd HH:mm:ss.sss", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, formatter = ExcelDefaultDataFormatTest.class, fillStyle = ExcelExportFillStyleEnum.VERTICAL, colWidth = -1, contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ)
	private Date createTime;
	@ExcelExportCell(titleName = "", titleModel = ExcelExportCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL, contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ, rowHeight = 51, colWidth = 430/*,mergeContentColNum = 1,mergeRowNum = 1,mergeTitleColNum = 1*/)
	private Map<String,String> map;
	@ExcelExportCell(titleName = "修改时间", formatPattern = "yyyy-MM-dd", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.HORIZONTAL,colWidth = 300)
	private Calendar updateTime;
	@ExcelExportCell(titleName = "logo", titleModel = ExcelExportCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelExportFillStyleEnum.VERTICAL, contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ, colIndex = 0, rowIndex = 0, rowHeight = 51, colWidth = 430)
	private byte[] logo;



}
