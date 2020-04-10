package com.github.excel.model;

import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:06 下午
 * @Description: 错误信息model
 */
@Data
public class ExcelReadErrorMsgInfo {
	public ExcelReadErrorMsgInfo(String errorMsg) {
		this.errorMsg = errorMsg;
	}

	public ExcelReadErrorMsgInfo(Integer sheetIndex, String sheetName, Integer rowIndex, Integer colIndex, String colPoint, String errorMsg) {
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
		this.rowIndex = rowIndex;
		this.colIndex = colIndex;
		this.colPoint = colPoint;
		this.errorMsg = errorMsg;
	}

	private Integer sheetIndex;
	private String sheetName;
	private Integer rowIndex;
	private Integer colIndex;
	private String colPoint;
	private String errorMsg;
}
