package com.github.export;

import com.github.excel.write.AbstractExcelStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:25 下午
 * @Description: Excel 自定义样式
 */
public class ExcelCustomStyle extends AbstractExcelStyle {


	public static final String STYLE_TITLE_TEST = "STYLE_TITLE_TEST";

	public static final String STYLE_CONTENT_TEST = "STYLE_CONTENT_TEST";

	public static final String FONT_SIZE16_BLOLD_RED_TEST = "FONT_SIZE16_BLOLD_RED_TEST";

	public static final String FONT_SIZE14_BLOLD_UNDERLINE_TEST = "FONT_SIZE14_BLOLD_UNDERLINE_TEST";

	private static final String DEFAULT_FONT_NAME1 = "黑体";


	public ExcelCustomStyle(Workbook workbook, Map<String, CellStyle> styleMap, Map<String, Font> fontMap) {
		super(workbook, styleMap, fontMap);
	}

	@Override
	public void addNewStyle() {
		CellStyle style = this.createStyle(STYLE_TITLE_TEST);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD_RED_TEST));
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLUE.getIndex());
		style.setBorderTop(BorderStyle.MEDIUM_DASHED);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());

		style = this.createStyle(STYLE_CONTENT_TEST);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(fontMap.get(FONT_SIZE14_BLOLD_UNDERLINE_TEST));

	}

	@Override
	public void addNewFont() {

		Font font = this.createFont(FONT_SIZE16_BLOLD_RED_TEST);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME1);
		font.setColor(IndexedColors.RED.index);

		font = this.createFont(FONT_SIZE14_BLOLD_UNDERLINE_TEST);
		font.setFontHeightInPoints((short) 14);
		font.setFontName(DEFAULT_FONT_NAME1);
		font.setBold(true);
		font.setUnderline(FontUnderline.SINGLE.getByteValue());


	}
}
