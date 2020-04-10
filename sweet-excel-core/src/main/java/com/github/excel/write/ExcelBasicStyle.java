package com.github.excel.write;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.helper.ExcelHelper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:17 下午
 * @Description: Excel 内置样式
 */
public class ExcelBasicStyle extends AbstractExcelStyle {

	public static final String STYLE_FOREGROUND_COLOR_YELLOW = "STYLE_FOREGROUND_COLOR_YELLOW";

	public static final String STYLE_AROUND_BORDER_READ = "STYLE_AROUND_BORDER_READ";

	public static final String STYLE_TITLE = "STYLE_TITLE";

	public static final String STYLE_LIST_TITLE = "STYLE_LIST_TITLE";

	public static final String STYLE_TITLE_RED_FONT = "STYLE_TITLE_RED_FONT";

	public static final String STYLE_CONTENT = "STYLE_CONTENT";

	public static final String STYLE_HLINK = "STYLE_HLINK";

	public static final String STYLE_ZEBRA_TITLE_ROW = "STYLE_ZEBRA_TITLE";

	public static final String STYLE_ZEBRA_ODD_ROW = "STYLE_ZEBRA_ODD_ROW";

	public static final String STYLE_ZEBRA_EVEN_ROW = "STYLE_ZEBRA_EVEN_ROW";

	public static final String STYLE_ZEBRA_ODD_ROW_DATE = "STYLE_ZEBRA_ODD_ROW_DATE";

	public static final String STYLE_ZEBRA_EVEN_ROW_DATE = "STYLE_ZEBRA_EVEN_ROW_DATE";

	public static final String STYLE_DATE_YYYYMMDDHHMMSS = "STYLE_DATE_YYYYMMDDHHMMSS";

	public static final String FONT_SIZE16_BLOLD = "FONT_SIZE16_BLOLD";

	public static final String FONT_SIZE16_BLOLD_RED = "FONT_SIZE16_BLOLD_RED";

	public static final String FONT_SIZE16_BLOLD_WHITE = "FONT_SIZE16_BLOLD_WHITE";

	public static final String FONT_SIZE14 = "FONT_SIZE14";

	public static final String FONT_SIZE14_BLOLD_UNDERLINE = "FONT_SIZE14_BLOLD_UNDERLINE";

	public static final String FONT_HLINK = "FONT_HLINK";

	private static final String DEFAULT_FONT_NAME = "黑体";


	public ExcelBasicStyle(Workbook workbook, Map<String, CellStyle> styleMap, Map<String, Font> fontMap) {
		super(workbook, styleMap, fontMap);
	}

	@Override
	public void addNewStyle() {
		CellStyle style = this.createStyle(STYLE_AROUND_BORDER_READ);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(IndexedColors.RED.index);
		style.setLeftBorderColor(IndexedColors.RED.index);
		style.setRightBorderColor(IndexedColors.RED.index);
		style.setTopBorderColor(IndexedColors.RED.index);

		style = this.createStyle(STYLE_FOREGROUND_COLOR_YELLOW);
		style.setFont(fontMap.get(FONT_SIZE14));
		style.setFillForegroundColor(IndexedColors.YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);

		style = this.createStyle(STYLE_TITLE);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD));

		style = this.createStyle(STYLE_LIST_TITLE);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD));

		style = this.createStyle(STYLE_TITLE_RED_FONT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD_RED));

		style = this.createStyle(STYLE_CONTENT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(fontMap.get(FONT_SIZE14));

		style = this.createStyle(STYLE_HLINK);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(fontMap.get(FONT_HLINK));

		style = this.createStyle(STYLE_DATE_YYYYMMDDHHMMSS);
		CreationHelper creationHelper = workbook.getCreationHelper();
		style.setDataFormat(creationHelper.createDataFormat().getFormat(ExcelConstant.DEFAULT_DATE_FORMAT));

		createTheme();
	}

	private void createTheme() {
		if (workbook instanceof HSSFWorkbook) {
			// ================= theme =================
			short titleRowBackGroundColorIndex = createHssfColor(workbook,  53,92,183,HSSFColor.HSSFColorPredefined.GREEN.getIndex());
			short borderColorIndex = createHssfColor(workbook,  125,152,210,HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex());
			short oDDColorIndex = createHssfColor(workbook,208,219,239,HSSFColor.HSSFColorPredefined.BLUE_GREY.getIndex());

			CellStyle style = this.createStyle(STYLE_ZEBRA_TITLE_ROW);
			style.setVerticalAlignment(VerticalAlignment.CENTER);
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setFillForegroundColor(titleRowBackGroundColorIndex);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);

			style.setBottomBorderColor(borderColorIndex);
			style.setTopBorderColor(borderColorIndex);
			style.setLeftBorderColor(borderColorIndex);
			style.setRightBorderColor(borderColorIndex);
			style.setFont(fontMap.get(FONT_SIZE16_BLOLD_WHITE));

			createHSSFContentTheme(borderColorIndex, oDDColorIndex,STYLE_ZEBRA_EVEN_ROW,STYLE_ZEBRA_ODD_ROW);
			createHSSFContentTheme(borderColorIndex, oDDColorIndex,STYLE_ZEBRA_EVEN_ROW_DATE,STYLE_ZEBRA_ODD_ROW_DATE);
			setDateFormat();
			// ================= theme end=================
		} else if (workbook instanceof XSSFWorkbook || workbook instanceof SXSSFWorkbook) {
			XSSFColor titleRowBackGroundColor = createXssfColor(workbook,  53,92,183);
			XSSFColor borderColor = createXssfColor(workbook,  125,152,210);
			XSSFColor oDDColor = createXssfColor(workbook,208,219,239);

			XSSFCellStyle style = (XSSFCellStyle)this.createStyle(STYLE_ZEBRA_TITLE_ROW);
			style.setVerticalAlignment(VerticalAlignment.CENTER);
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setFillForegroundColor(titleRowBackGroundColor);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);

			style.setTopBorderColor(borderColor);
			style.setLeftBorderColor(borderColor);
			style.setRightBorderColor(borderColor);
			style.setBottomBorderColor(borderColor);

			style.setFont(fontMap.get(FONT_SIZE16_BLOLD_WHITE));

			createXSSFContentTheme(borderColor, oDDColor,STYLE_ZEBRA_EVEN_ROW,STYLE_ZEBRA_ODD_ROW);
			createXSSFContentTheme(borderColor, oDDColor,STYLE_ZEBRA_EVEN_ROW_DATE,STYLE_ZEBRA_ODD_ROW_DATE);
			setDateFormat();
		}
	}

	private void setDateFormat() {
		short format = styleMap.get(STYLE_DATE_YYYYMMDDHHMMSS).getDataFormat();
		CellStyle cellStyle = styleMap.get(STYLE_ZEBRA_EVEN_ROW_DATE);
		cellStyle.setDataFormat(format);
		cellStyle = styleMap.get(STYLE_ZEBRA_ODD_ROW_DATE);
		cellStyle.setDataFormat(format);
	}
	private void createXSSFContentTheme(XSSFColor borderColor, XSSFColor oDDColor,String styleEventRow,String styleOddRow) {
		XSSFCellStyle style;
		style = (XSSFCellStyle)this.createStyle(styleEventRow);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColor);
		style.setTopBorderColor(borderColor);
		style.setLeftBorderColor(borderColor);
		style.setRightBorderColor(borderColor);
		style.setFont(fontMap.get(FONT_SIZE14));

		style = (XSSFCellStyle)this.createStyle(styleOddRow);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(oDDColor);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColor);
		style.setTopBorderColor(borderColor);
		style.setLeftBorderColor(borderColor);
		style.setRightBorderColor(borderColor);
		style.setFont(fontMap.get(FONT_SIZE14));
	}

	private void createHSSFContentTheme(short borderColorIndex, short oDDColorIndex,String styleEventRow,String styleOddRow) {
		CellStyle style;
		style = this.createStyle(styleEventRow);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColorIndex);
		style.setTopBorderColor(borderColorIndex);
		style.setLeftBorderColor(borderColorIndex);
		style.setRightBorderColor(borderColorIndex);
		style.setFont(fontMap.get(FONT_SIZE14));

		style = this.createStyle(styleOddRow);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(oDDColorIndex);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColorIndex);
		style.setTopBorderColor(borderColorIndex);
		style.setLeftBorderColor(borderColorIndex);
		style.setRightBorderColor(borderColorIndex);
		style.setFont(fontMap.get(FONT_SIZE14));
	}

	private short createHssfColor(Workbook workbook,int r,int g,int b,short colorIndex) {
		HSSFColor hssfColor = ExcelHelper.setCustomColor((HSSFWorkbook) workbook, (byte) r, (byte)g, (byte)b,colorIndex);
		return hssfColor.getIndex();
	}

	private XSSFColor createXssfColor(Workbook workbook,int r,int g,int b) {
		String indexKey = String.valueOf(r) + String.valueOf(g) + String.valueOf(b);
		XSSFColor xssfColor = colorMap.get(indexKey);
		if (Objects.nonNull(xssfColor)) {
			return xssfColor ;
		}
		XSSFColor color = new XSSFColor(new java.awt.Color(r, g, b));
		colorMap.put(indexKey, color);
		return color;

	}

	@Override
	public void addNewFont() {
		Font font = this.createFont(FONT_SIZE16_BLOLD);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME);

		font = this.createFont(FONT_SIZE16_BLOLD_RED);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME);
		font.setColor(IndexedColors.RED.index);

		font = this.createFont(FONT_SIZE16_BLOLD_WHITE);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME);
		font.setColor(IndexedColors.WHITE.index);

		font = this.createFont(FONT_SIZE14);
		font.setFontHeightInPoints((short) 14);
		font.setFontName(DEFAULT_FONT_NAME);

		font = this.createFont(FONT_SIZE14_BLOLD_UNDERLINE);
		font.setFontHeightInPoints((short) 14);
		font.setFontName(DEFAULT_FONT_NAME);
		font.setBold(true);
		font.setUnderline(FontUnderline.SINGLE.getByteValue());

		font = this.createFont(FONT_HLINK);
		font.setUnderline(Font.U_SINGLE);
		font.setColor(IndexedColors.BLUE.getIndex());


	}
}
