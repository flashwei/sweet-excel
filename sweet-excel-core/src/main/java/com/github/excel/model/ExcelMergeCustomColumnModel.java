package com.github.excel.model;

import lombok.Data;

import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:05 下午
 * @Description: 导出到单个单元格model
 */
@Data
public class ExcelMergeCustomColumnModel extends BaseExcelColumnModel {
	private int firstRow;
	private int lastRow;
	private int firstColumn;
	private int lastColumn;

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (!(o instanceof ExcelMergeCustomColumnModel)) return false;
		if (!super.equals(o)) return false;
		ExcelMergeCustomColumnModel that = (ExcelMergeCustomColumnModel) o;
		return getFirstRow() == that.getFirstRow() && getLastRow() == that.getLastRow() && getFirstColumn() == that.getFirstColumn() && getLastColumn() == that.getLastColumn();
	}

	@Override
	public int hashCode() {
		return Objects.hash(super.hashCode(), getFirstRow(), getLastRow(), getFirstColumn(), getLastColumn());
	}
}
