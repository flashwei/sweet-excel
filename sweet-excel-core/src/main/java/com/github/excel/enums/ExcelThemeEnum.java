package com.github.excel.enums;

import com.github.excel.write.ExcelBasicStyle;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:00 下午
 * @Description: excel 主题样式
 */
public enum ExcelThemeEnum {
	NONE(null, null, null, null, null, null,null,null),
	ZEBRA(ExcelBasicStyle.STYLE_ZEBRA_TITLE_ROW, ExcelBasicStyle.STYLE_ZEBRA_ODD_ROW, ExcelBasicStyle.STYLE_ZEBRA_EVEN_ROW, (short)40, (short)32, (short)165,ExcelBasicStyle.STYLE_ZEBRA_ODD_ROW_DATE, ExcelBasicStyle.STYLE_ZEBRA_EVEN_ROW_DATE),
	;
	/**
	 * 标题行样式名称
	 */
	private String titleRowStyleName;
	/**
	 * 奇数行样式名称
	 */
	private String oddRowStyleName;
	/**
	 * 偶数行样式名称
	 */
	private String evenRowStyleName;
	/**
	 * 奇数行样式名称
	 */
	private String oddRowStyleDateName;
	/**
	 * 偶数行样式名称
	 */
	private String evenRowStyleDateName;
	/**
	 * 标题默认高度
	 */
	private Short titleRowHeight;
	/**
	 * 内容默认高度
	 */
	private Short contentRowHeight;
	/**
	 * 单元格默认宽度
	 */
	private Short colWidth;

	ExcelThemeEnum(String titleRowStyleName, String oddRowStyleName, String evenRowStyleName, Short titleRowHeight, Short contentRowHeight, Short colWidth,String oddRowStyleDateName,String evenRowStyleDateName) {
		this.titleRowStyleName = titleRowStyleName;
		this.oddRowStyleName = oddRowStyleName;
		this.evenRowStyleName = evenRowStyleName;
		this.titleRowHeight = titleRowHeight;
		this.contentRowHeight = contentRowHeight;
		this.colWidth = colWidth;
		this.oddRowStyleDateName = oddRowStyleDateName;
		this.evenRowStyleDateName = evenRowStyleDateName;
	}

	public String getTitleRowStyleName() {
		return titleRowStyleName;
	}

	public String getOddRowStyleName() {
		return oddRowStyleName;
	}

	public String getEvenRowStyleName() {
		return evenRowStyleName;
	}

	public Short getTitleRowHeight() {
		return titleRowHeight;
	}

	public Short getContentRowHeight() {
		return contentRowHeight;
	}

	public Short getColWidth() {
		return colWidth;
	}

	public String getOddRowStyleDateName() {
		return oddRowStyleDateName;
	}

	public String getEvenRowStyleDateName() {
		return evenRowStyleDateName;
	}
}
