package com.github.excel.annotation;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelExportCellTitleModelEnum;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.enums.ExcelExportListFillTypeEnum;
import com.github.excel.enums.ExcelExportScopeEnum;
import com.github.excel.enums.ExcelThemeEnum;

import java.lang.annotation.*;


@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelExport {
	/**
	 * 命名空间
	 */
	String nameSpace() default ExcelConstant.NULL_STR;

	/**
	 * 作用域
	 * @see ExcelExportScopeEnum
	 */
	ExcelExportScopeEnum scope() default ExcelExportScopeEnum.ALL_SHEET;

	/**
	 * 填充标题
	 */
	boolean fillTitle() default true;

	/**
	 * 冻结标题
	 * @return
	 */
	boolean freezeTitle() default false;

	/**
	 * 行号
	 * @return
	 */
	int rowIndex() default ExcelConstant.ZERO_SHORT;

	/**
	 * 列号
	 * @return
	 */
	int colIndex() default ExcelConstant.ZERO_SHORT;

	/**
	 * 填充样式
	 * @see ExcelExportFillStyleEnum
	 */
	ExcelExportFillStyleEnum fillStyle() default ExcelExportFillStyleEnum.VERTICAL;

	/**
	 * 标题样式名
	 * @return
	 */
	String titleStyleName() default ExcelConstant.NULL_STR;

	/**
	 * 内容样式名
	 * @return
	 */
	String contentStyleName() default ExcelConstant.NULL_STR;

	/**
	 * 列表填充类型
	 * @see ExcelExportListFillTypeEnum
	 */
	ExcelExportListFillTypeEnum fillType() default ExcelExportListFillTypeEnum.COVER;

	/**
	 * 标题行合并数量
	 * @return
	 */
	int mergeTitleRowNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 标题列合并数量
	 * @return
	 */
	int mergeTitleColNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 内容行合并数量
	 * @return
	 */
	int mergeContentRowNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 内容列合并数量
	 * @return
	 */
	int mergeContentColNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 标题填充样式
	 * @see ExcelExportCellTitleModelEnum
	 */
	ExcelExportCellTitleModelEnum titleModel() default ExcelExportCellTitleModelEnum.DEFAULT;

	/**
	 * theme
	 * @return
	 */
	ExcelThemeEnum theme() default ExcelThemeEnum.NONE;

	/**
	 * 是否填充自增序列号
	 * @return
	 */
	boolean incrementSequenceNo() default false;

	/**
	 * 是否填充自增序列号
	 * @return
	 */
	String incrementSequenceTitle() default ExcelConstant.INCREMENT_SEQUENCE_NO_TITLE_NAME;
}
