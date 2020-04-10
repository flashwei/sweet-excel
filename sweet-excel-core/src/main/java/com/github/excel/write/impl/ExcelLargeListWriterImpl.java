package com.github.excel.write.impl;

import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.boot.WorkbookCachePool;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.util.ExcelUtil;
import com.github.excel.write.AbstractExcelStyle;
import com.github.excel.write.BaseExcelWriter;
import com.github.excel.write.ExcelWriterDataFormat;
import com.google.common.base.Throwables;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelWriteException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelLargeListWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicLong;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:16 下午
 * @Description: excel 导出大list数据实现
 */
@Slf4j
public class ExcelLargeListWriterImpl extends BaseExcelWriter implements ExcelLargeListWriter {

	private final SXSSFWorkbook workbook;

	private final ThreadLocal<WorkbookCachePool.WorkbookCacheModel> sxssfWorkbookThreadLocal ;

	private String sheetName;

	private boolean fillTitle = true;

	private int sheetNo = ExcelConstant.ONE_INT;

	private int rowIndex = ExcelConstant.ZERO_SHORT;

	private int sheetRowMaxCount ;

	public ExcelLargeListWriterImpl(String sheetName) {
		sxssfWorkbookThreadLocal = WorkbookCachePool.getSxssfWorkbookThreadLocal();
		WorkbookCachePool.WorkbookCacheModel workbookCacheModel = sxssfWorkbookThreadLocal.get();
		workbook = (SXSSFWorkbook) workbookCacheModel.getWorkbook();
		this.styleLocal.set(workbookCacheModel.getStyleMap());
		this.fontLocal.set(workbookCacheModel.getFontMap());
		this.sheetRowMaxCount = ExcelConstant.INT_1000000 ;
		this.sheetName = sheetName;
	}

	public ExcelLargeListWriterImpl(String sheetName , int sheetRowMaxCount) {
		this(sheetName);
		if (sheetRowMaxCount > ExcelConstant.ZERO_SHORT && sheetRowMaxCount <= ExcelConstant.INT_1000000) {
			this.sheetRowMaxCount = sheetRowMaxCount;
		}else{
			this.sheetRowMaxCount = ExcelConstant.INT_1000000 ;
		}
	}

	public ExcelLargeListWriterImpl(String sheetName , int sheetRowMaxCount,Class<? extends ExcelBaseModel> listCla) {
		this(sheetName, sheetRowMaxCount);
		this.listCla = listCla;
	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList,String[] excludeFields) {

		if (Objects.isNull(modelList) || modelList.isEmpty()) {
			throw new ExcelWriteException("modelList can't be null");
		}

		Map<String, String> excludeFieldMap = new HashMap<>();
		if(Objects.nonNull(excludeFields)) {
			for (String excludeField : excludeFields) {
				excludeFieldMap.put(excludeField, ExcelConstant.NULL_STR);
			}
		}
		Class<? extends ExcelBaseModel> modelClass = modelList.get(ExcelConstant.ZERO_SHORT).getClass();
		ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(modelClass);
		CreationHelper creationHelper = workbook.getCreationHelper();
		SXSSFSheet sheet;
		List<ExcelCacheModel.ExcelCacheFieldModel> cacheFieldModelList = new ArrayList<>();
		for (ExcelCacheModel.ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
			if (Objects.nonNull(excludeFieldMap.get(fieldModel.getFieldName()))) {
				continue;
			}
			cacheFieldModelList.add(fieldModel);
		}
		// 填充标题
		if (fillTitle) {
			rowIndex = ExcelConstant.ZERO_SHORT;
			sheet = workbook.createSheet(sheetName + ExcelConstant.SHORT_TERM + sheetNo);
			SXSSFRow row = sheet.createRow(rowIndex);
			int colIndex = ExcelConstant.ZERO_SHORT;
			for (ExcelCacheModel.ExcelCacheFieldModel fieldModel : cacheFieldModelList) {
				ExcelExportCell exportCell = fieldModel.getExportCell();
				String titleName = cacheModel.getExcelExport().incrementSequenceTitle();
				if (Objects.nonNull(exportCell)) {
					titleName = exportCell.titleName();
				}
				SXSSFCell cell = row.createCell(colIndex);
				ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
				ExcelHelper.setRowHeight(row,fieldModel.getTitleRowHeight());
				setCellValueAndStyle(cell, titleName, fieldModel.getTitleStyleName(), null, creationHelper);
//				cell.setCellValue(fieldModel.getExportCell().titleName());
				++colIndex;
			}
			fillTitle = false;
			++rowIndex;
		} else {
			sheet = workbook.getSheet(sheetName + ExcelConstant.SHORT_TERM + sheetNo);
		}
		// 填充内容
		int i = ExcelConstant.ZERO_SHORT;
		for (T model : modelList) {
			int colIndex = ExcelConstant.ZERO_SHORT;
			i++;
			for (ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel : cacheFieldModelList) {
				ExcelExportCell exportCell = cacheFieldModel.getExportCell();
				Object value ;
				short colWidth = ExcelConstant.MINUS_TWO_SHORT;
				String linkName = ExcelConstant.NULL_STR;
				if(Objects.nonNull(exportCell)) {
					colWidth = exportCell.colWidth();
					linkName = exportCell.linkName();
					try {
						value = cacheFieldModel.getGetMethod().invoke(model);
						ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());
						value = formatValue(value, exportCell.formatPattern(), formatter);
					} catch (IllegalAccessException e) {
						log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
						throw new ExcelWriteException(e.getMessage());
					} catch (InvocationTargetException e) {
						log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
						throw new ExcelWriteException(e.getMessage());
					}
				}else{
					AtomicLong incrementSeq = incrementSeqMap.get(modelClass);
					if (incrementSeq == null) {
						incrementSeq = new AtomicLong(ExcelConstant.ZERO_SHORT);
						incrementSeqMap.put(modelClass, incrementSeq);
					}
					value = incrementSeq.incrementAndGet();
				}
				String styleName = i % ExcelConstant.TOW_INT == ExcelConstant.ZERO_SHORT ? cacheFieldModel.getEvenRowStyleName() : cacheFieldModel.getContentStyleName();

				ExcelHelper.setColWidth(sheet, colIndex, colWidth);

				Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
				Cell cell = row.createCell(colIndex);
				ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				if (Objects.nonNull(value) && value instanceof byte[]) {
					createPicture(workbook, sheet, cell, (byte[]) value, styleName);
				} else {
					setCellValueAndStyle(cell, value, styleName, linkName, creationHelper);
				}
				colIndex++;
			}
			colIndex = ExcelConstant.ZERO_SHORT;
			rowIndex++;
		}
		// 到达sheetRowMaxCount创建新的sheet
		if (rowIndex >= sheetRowMaxCount) {
			sheetNo++;
			fillTitle = true;
		}
	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList) {
		process(modelList,null);
	}

	@Override
	public void setNoneDataTips(boolean noneDataTips) {
		this.noneDataTips = noneDataTips;
	}

	@Override
	public void export(OutputStream outputStream) {
		try {
			CreationHelper creationHelper = workbook.getCreationHelper();
			addNoResultData(workbook, creationHelper);
			workbook.write(outputStream);
		} catch (IOException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}finally {
			try {
				outputStream.flush();
				outputStream.close();
			} catch (IOException e) {
				log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
				throw new ExcelWriteException(e.getMessage());
			}
			workbook.dispose();
			fontLocal.remove();
			styleLocal.remove();
			sxssfWorkbookThreadLocal.remove();
		}
	}

	@Override
	public void export(HttpServletRequest request, HttpServletResponse response, String fileName) {
		try {
			CreationHelper creationHelper = workbook.getCreationHelper();
			addNoResultData(workbook, creationHelper);
			ExcelUtil.setResponseHeader(request, response, fileName, ExcelConstant.XLSX_STR);
			OutputStream outputStream = response.getOutputStream();
			workbook.write(outputStream);
		} catch (IOException e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}finally {
			workbook.dispose();
			fontLocal.remove();
			styleLocal.remove();
			sxssfWorkbookThreadLocal.remove();
		}
	}

	@Override
	public void addStyle(Class<? extends AbstractExcelStyle> styleClass) {
		initStyle(workbook,styleClass);
	}

	@Override
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	@Override
	public void close() {

	}
}
