package com.github.excel.model;

import com.github.excel.write.ExcelDefaultWriterDataFormat;
import com.github.excel.write.ExcelWriterDataFormat;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelExportColumnFillTypeEnum;
import lombok.Data;

import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:02 下午
 * @Description: 基础导出模型
 */
@Data
public class BaseExcelColumnModel{
	private String sheetName;
	private Object value;
	private String styleName;
	private short colWidth = ExcelConstant.MINUS_TWO_SHORT ;
	private short rowHeight = ExcelConstant.MINUS_TWO_SHORT ;
	private String formatPattern;
	private Class<? extends ExcelWriterDataFormat> dataFormat = ExcelDefaultWriterDataFormat.class;
	private ExcelExportColumnFillTypeEnum fillType = ExcelExportColumnFillTypeEnum.APPEND;

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (!(o instanceof BaseExcelColumnModel)) return false;
		BaseExcelColumnModel that = (BaseExcelColumnModel) o;
		return colWidth == that.colWidth && rowHeight == that.rowHeight && Objects.equals(sheetName, that.sheetName) && Objects.equals(value, that.value) && Objects.equals(styleName, that.styleName) && Objects.equals(formatPattern, that.formatPattern) && Objects.equals(dataFormat, that.dataFormat) && fillType == that.fillType;
	}

	@Override
	public int hashCode() {
		return Objects.hash(sheetName, value, styleName, colWidth, rowHeight, formatPattern, dataFormat, fillType);
	}
}
