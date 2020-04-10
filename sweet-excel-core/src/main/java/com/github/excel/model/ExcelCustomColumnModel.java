package com.github.excel.model;

import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:04 下午
 * @Description: 导出到单个单元格model
 */
@Data
public class ExcelCustomColumnModel extends BaseExcelColumnModel {
	private int rowIndex;
	private int colIndex;

	@Override
	public boolean equals(Object obj) {
		return (this == obj);
	}

}
