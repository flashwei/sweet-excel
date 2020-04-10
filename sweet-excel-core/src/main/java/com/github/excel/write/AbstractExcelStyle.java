package com.github.excel.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:16 下午
 * @Description: excel样式抽象类
 */
public abstract class AbstractExcelStyle {

	protected Workbook workbook;
	protected Map<String, CellStyle> styleMap;
	protected Map<String, Font> fontMap;
	protected Map<String, XSSFColor> colorMap = new HashMap<>();

	public AbstractExcelStyle(Workbook workbook, Map<String, CellStyle> styleMap, Map<String, Font> fontMap) {
		this.workbook = workbook;
		this.styleMap = styleMap;
		this.fontMap = fontMap;
	}

	/**
	 * 添加新样式
	 */
	public abstract void addNewStyle();

	/**
	 * 添加新字体
	 */
	public abstract void addNewFont();

	/**
	 * 创建样式
	 *
	 * @param name 名称
	 * @return CellStyle
	 */
	public CellStyle createStyle(String name) {
		CellStyle style = styleMap.get(name);
		if (Objects.nonNull(style)) {
			return style;
		}
		style = workbook.createCellStyle();
		styleMap.put(name, style);
		return style;
	}

	/**
	 * 创建字体
	 *
	 * @param name 名称
	 * @return Font
	 */
	public Font createFont(String name) {
		Font font = fontMap.get(name);
		if (Objects.nonNull(font)) {
			return font;
		}
		font = workbook.createFont();
		fontMap.put(name, font);
		return font;
	}
}

