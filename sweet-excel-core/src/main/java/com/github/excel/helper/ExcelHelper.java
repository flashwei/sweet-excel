package com.github.excel.helper;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.exception.ExcelWriteException;
import com.github.excel.model.ExcelRichTextModel;
import com.github.excel.util.StringUtil;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.Calendar;
import java.util.Date;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:01 下午
 * @Description: excel 帮助类
 */
public class ExcelHelper {
	/**
	 * 创建富文本
	 *
	 * @param creationHelper    创建
	 * @param text              文本
	 * @param richTextModelList 富文本属性List
	 * @return RichTextString
	 */
	public static RichTextString createRichText(CreationHelper creationHelper, String text, ExcelRichTextModel... richTextModelList) {
		RichTextString richTextString = creationHelper.createRichTextString(text);
		for (ExcelRichTextModel textModel : richTextModelList) {
			if (null == textModel) {
				continue;
			}
			richTextString.applyFont(textModel.getStartIndex(), textModel.getEndIndex(), textModel.getFont());
		}
		return richTextString;
	}

	/**
	 * 获取超链接
	 *
	 * @param createHelper helper
	 * @param link         连接
	 * @return Hyperlink
	 */
	public static Hyperlink getHyperlink(CreationHelper createHelper, String link) {
		Hyperlink hyperlink = null;
		if (link.startsWith(ExcelConstant.HTTPS_PROTOCOL) || link.startsWith(ExcelConstant.HTTP_PROTOCOL)) {
			hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
		} else if (link.startsWith(ExcelConstant.EMAIL_PROTOCOL)) {
			hyperlink = createHelper.createHyperlink(HyperlinkType.EMAIL);
		} else if (link.startsWith(ExcelConstant.FILE_PROTOCOL)) {
			hyperlink = createHelper.createHyperlink(HyperlinkType.FILE);
		}
		return hyperlink;
	}

	/**
	 * 创建链接
	 *
	 * @param cell
	 * @param value
	 * @param linkName
	 * @param createHelper
	 * @return
	 */
	public static Object createHyperlink(Cell cell, Object value, String linkName, CreationHelper createHelper) {
		if (StringUtil.notEmpty(linkName)) {
			String address = value.toString();
			Hyperlink hyperlink = getHyperlink(createHelper, address);
			if (null != hyperlink) {
				hyperlink.setAddress(address);
				cell.setHyperlink(hyperlink);
			}
			if (!linkName.equals(ExcelConstant.MINUS_ONE_STR)) {
				value = linkName;
			}
		}
		return value;
	}

	/**
	 * 设置单元格值
	 *
	 * @param cell
	 * @param value
	 */
	public static void setCellValue(Cell cell, Object value) {
		// 设置内容
		if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Number) {
			Number number = (Number) value;
			cell.setCellValue(number.doubleValue());
		} else if (value instanceof Boolean) {
			cell.setCellValue((boolean) value);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar) value);
		} else if (value instanceof RichTextString) {
			cell.setCellValue((RichTextString) value);
		} else {
			cell.setCellValue(value.toString());
		}
	}

	/**
	 * 根据名称获取sheet，不存在则创建
	 *
	 * @param workbook workbook
	 * @param name     名称
	 * @return Sheet
	 */
	public static Sheet getSheetOrCreate(Workbook workbook, String name) {
		Sheet sheet = workbook.getSheet(name);
		if (Objects.isNull(sheet)) {
			sheet = workbook.createSheet(name);
		}
		return sheet;
	}

	/**
	 * 获取行，不存在则创建
	 *
	 * @param sheet    sheet对象
	 * @param rowIndex 行坐标
	 * @return Row
	 */
	public static Row getRowOrCreate(Sheet sheet, int rowIndex) {
		synchronized (sheet) {
			Row row = sheet.getRow(rowIndex);
			if (Objects.isNull(row)) {
				row = sheet.createRow(rowIndex);
			}
			return row;
		}
	}

	/**
	 * 获取单元格，不存在则创建
	 *
	 * @param row       行
	 * @param cellIndex 单元格坐标
	 * @return Cell
	 */
	public static Cell getCellOrCreate(Row row, int cellIndex) {
		Cell cell = row.getCell(cellIndex);
		if (Objects.isNull(cell)) {
			cell = row.createCell(cellIndex);
		}
		return cell;
	}

	/**
	 * 设置行高
	 *
	 * @param row       行
	 * @param rowHeight 行高
	 */
	public static void setRowHeight(Row row, short rowHeight) {
		if (ExcelConstant.MINUS_TWO_SHORT != rowHeight) {
			row.setHeightInPoints(rowHeight);
		}
	}

	/**
	 * 设置列宽
	 *
	 * @param sheet     sheet
	 * @param cellIndex 列坐标
	 * @param width     宽度
	 */
	public static void setColWidth(Sheet sheet, int cellIndex, short width) {
		if (ExcelConstant.MINUS_TWO_SHORT != width) {
			int widthPixel = (int) ExcelConstant.PIXEL_RATE * width;
			if (ExcelConstant.MINUS_ONE_SHORT == width) {
				sheet.autoSizeColumn(cellIndex);
			} else {
				sheet.setColumnWidth(cellIndex, widthPixel);
			}
		}
	}

	/**
	 * 移动row
	 *
	 * @param sheet
	 * @param rowIndex
	 * @return
	 */
	public static void shiftRows(Sheet sheet, Integer rowIndex, int shiftRowSize) {
		if (sheet.getRow(rowIndex) != null) {
			int lastRowNo = sheet.getLastRowNum();
			sheet.shiftRows(rowIndex, lastRowNo, shiftRowSize);
		}
	}

	/**
	 * 获取总行数
	 *
	 * @param workbook
	 * @param sheetName
	 * @return
	 */
	public static int getRowNum(Workbook workbook, String sheetName) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (Objects.isNull(sheet)) {
			return ExcelConstant.ZERO_SHORT;
		}
		return sheet.getLastRowNum() + ExcelConstant.ONE_INT;
	}

	/**
	 * 获取总列数
	 *
	 * @param workbook
	 * @param sheetName
	 * @return
	 */
	public static int getCellNum(Workbook workbook, String sheetName, int rowIndex) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (Objects.isNull(sheet)) {
			return ExcelConstant.ZERO_SHORT;
		}
		Row row = sheet.getRow(rowIndex);
		if (Objects.isNull(row)) {
			return ExcelConstant.ZERO_SHORT;
		}
		return row.getPhysicalNumberOfCells();
	}

	/**
	 * 添加下拉框数据验证
	 *
	 * @param sheet
	 * @param addressList
	 * @param DropDownArray
	 */
	public static void addDropDownValidation(Sheet sheet, CellRangeAddressList addressList, String[] DropDownArray) {
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		DataValidationConstraint dvConstraint = dvHelper.createExplicitListConstraint(DropDownArray);
		DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
		if (validation instanceof XSSFDataValidation) {
			validation.setSuppressDropDownArrow(true);
			validation.setShowErrorBox(true);
		} else {
			validation.setSuppressDropDownArrow(false);
		}
		validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
		validation.createErrorBox(ExcelErrorMsgConstant.ERROR_DROP_DOWN_TITLE, ExcelErrorMsgConstant.ERROR_DROP_DOWN_MSG);
		sheet.addValidationData(validation);
	}

	/**
	 * 添加范围验证
	 *
	 * @param sheet
	 * @param addressList
	 * @param startNumStr
	 * @param endNumStr
	 */
	public static void addRangeValidation(Sheet sheet, CellRangeAddressList addressList, String startNumStr, String endNumStr) {
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		DataValidationConstraint dvConstraint = dvHelper.createNumericConstraint(DVConstraint.ValidationType.INTEGER, DVConstraint.OperatorType.BETWEEN, startNumStr, endNumStr);
		DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
		if (validation instanceof XSSFDataValidation) {
			validation.setSuppressDropDownArrow(true);
			validation.setShowErrorBox(true);
		} else {
			validation.setSuppressDropDownArrow(false);
		}
		validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
		validation.createErrorBox(ExcelErrorMsgConstant.ERROR_DROP_DOWN_TITLE, String.format(ExcelErrorMsgConstant.ERROR_RANGE_MSG, startNumStr, endNumStr));
		sheet.addValidationData(validation);
	}

	/**
	 * 创建命名
	 *
	 * @param wb
	 * @param namename
	 * @param formula
	 */
	public static void addRefersToFormula(Workbook wb, String namename, String formula) {
		Name name = wb.createName();
		name.setNameName(namename);
		name.setRefersToFormula(formula);
	}

	/**
	 * 创建批注
	 *
	 * @param wb
	 * @param sheet
	 * @param rowIndex
	 * @param colIndex
	 */
	public static void createComment(Workbook wb, Sheet sheet, int rowIndex, int colIndex, String author, String text, Font font) {
		if (Objects.isNull(wb)) {
			throw new ExcelWriteException("Workbook can't be null");
		}
		if (Objects.isNull(sheet)) {
			throw new ExcelWriteException("sheet can't be null");
		}
		CreationHelper creationHelper = wb.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();

		Row row = getRowOrCreate(sheet, rowIndex);
		Cell cell = getCellOrCreate(row, colIndex);

		// When the comment box is visible, have it show in a 1x3 space
		ClientAnchor anchor = creationHelper.createClientAnchor();
		anchor.setCol1(cell.getColumnIndex());
		anchor.setCol2(cell.getColumnIndex() + ExcelConstant.THREE_INT);
		anchor.setRow1(row.getRowNum());
		anchor.setRow2(row.getRowNum() + ExcelConstant.THREE_INT);

		// Create the comment and set the text+author
		Comment comment = drawing.createCellComment(anchor);
		RichTextString str = creationHelper.createRichTextString(text);
		if (Objects.nonNull(font)) {
			str.applyFont(font);
		}
		comment.setString(str);
		comment.setAuthor(author);

		// Assign the comment to the cell
		cell.setCellComment(comment);
	}

	/**
	 * 设置sheet 缩放率
	 * @param wb workbook
	 * @param sheetName sheetName
	 * @param scale scale
	 */
	public static void setSheetZoom(Workbook wb, String sheetName,int scale) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			sheet.setZoom(scale);
		}
	}

	/**
	 * 设置打印区域
	 * @param wb
	 * @param sheetIndex
	 * @param startColIndex
	 * @param endColIndex
	 * @param startRowIndex
	 * @param endRowIndex
	 */
	public static void setPrintArea(Workbook wb, int sheetIndex,int startColIndex,int endColIndex,int startRowIndex,int endRowIndex) {
		wb.setPrintArea(sheetIndex,startColIndex,endColIndex,startRowIndex,endRowIndex);
	}

	/**
	 * 设置脚部页码
	 * @param wb
	 * @param sheetName
	 */
	public static void setFooterNumberByDefault(Workbook wb, String sheetName) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			Footer footer = sheet.getFooter();
			footer.setCenter(String.format(ExcelConstant.STRING_DEFAULT_FOOTER_TEXT, HeaderFooter.page(),HeaderFooter.numPages()));
		}
	}

	/**
	 * 设置脚部页码
	 * @param wb
	 * @param sheetName
	 */
	public static void setFooterNumber(Workbook wb, String sheetName,String formatPattern) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			Footer footer = sheet.getFooter();
			footer.setCenter(String.format(formatPattern, HeaderFooter.page(),HeaderFooter.numPages()));
		}
	}

	/**
	 * 创建拆分窗格
	 * @param wb workbook
	 * @param sheetName sheetName
	 * @param xSplitPos x 轴坐标，等同于像素
	 * @param ySplitPos y 坐标，等同于像素
	 * @param leftmostColumn 左边单元格数量
	 * @param topRow 上面单元格数量
	 */
	public static void createSplitPane(Workbook wb, String sheetName,int xSplitPos, int ySplitPos, int leftmostColumn, int topRow) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			sheet.createSplitPane(xSplitPos,ySplitPos,leftmostColumn,topRow,Sheet.PANE_LOWER_LEFT);
		}
	}

	/**
	 * 自定义颜色
	 * @param workbook
	 * @param r
	 * @param g
	 * @param b
	 * @return
	 */
	public static HSSFColor setCustomColor(HSSFWorkbook workbook, byte r, byte g, byte b,short colorIndex){
		HSSFPalette palette = workbook.getCustomPalette();
		HSSFColor hssfColor;
		try {
			hssfColor= palette.findColor(r, g, b);
			if (hssfColor == null ){
				palette.setColorAtIndex(colorIndex, r, g,b);
				hssfColor = palette.getColor(colorIndex);
			}
		} catch (Exception e) {
			throw new ExcelWriteException(e.getMessage());
		}
		return hssfColor;
	}

}
