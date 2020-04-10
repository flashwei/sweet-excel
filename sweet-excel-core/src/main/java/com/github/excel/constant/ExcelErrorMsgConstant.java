package com.github.excel.constant;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 6:59 下午
 * @Description: excel 错误常量
 */
public class ExcelErrorMsgConstant {

	public static final String ERROR_LOAD = "Unkown model. Please check model is loaded";
	public static final String ERROR_SHEET_NAME = "Sheet name can't be null";
	public static final String ERROR_COLUMN_POINT = "Column point error";
	public static final String ERROR_COLUMN_VALUE_NULL = "Column value can't be null";
	public static final String ERROR_DROP_DOWN_TITLE = "输入错误";
	public static final String ERROR_DROP_DOWN_MSG = "请选择下拉框里的值";
	public static final String ERROR_RANGE_MSG = "请输入%s~%s之间的值";
	public static final String ERROR_DATA_FORMAT_READ_MSG = "sheet:%s，单元格:%s，[%s]数据格式格式错误!";
	public static final String ERROR_DATA_NULL_READ_MSG = "sheet:%s，单元格:%s，[%s]数据不能为空!";
	public static final String ERROR_DATA_INVOKE_READ_MSG = "sheet:%s，单元格:%s，[%s]数据格式不正确!";
	public static final String ERROR_NOT_FOUND_SHEET = "sheet:%s，不存在";
	public static final String ERROR_NOT_FOUND_SHEET_ROW = "sheet:%s，第%s行不存在";
	public static final String ERROR_NOT_FOUND_SHEET_ROW_COL = "sheet:%s，第%s行第%s列不存在";
	public static final String ERROR_NOT_MATCH_COL_CONTENT = "sheet:%s，单元格:%s，内容不匹配，请修改为\"%s\"";
	public static final String ERROR_EXPORT_NOT_FOUND_DATA = "没有找到需要导出的数据，请修改筛选条件后再试！";
	public static final String ERROR_IMPORT_FILE_SUFFIX = "导入文件后缀名错误，请转换为%s后再试!";
}
