package com.github.excel.write;

import com.github.excel.model.NumberScopeModel;
import com.google.common.base.Strings;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.boot.WorkbookCachePool;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.enums.ExcelExportCellTitleModelEnum;
import com.github.excel.enums.ExcelExportColumnFillTypeEnum;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.enums.ExcelExportListFillTypeEnum;
import com.github.excel.exception.ExcelWriteException;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.BaseExcelColumnModel;
import com.github.excel.model.ComboBoxModel;
import com.github.excel.model.CommentModel;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.model.ExcelCustomColumnModel;
import com.github.excel.model.ExcelExportModel;
import com.github.excel.model.ExcelExpressionModel;
import com.github.excel.model.ExcelMergeCustomColumnModel;
import com.github.excel.model.ExcelRichTextModel;
import com.github.excel.model.ExcelTemplateCacheModel;
import com.github.excel.model.ExcelTemplateTitleModel;
import com.github.excel.util.ReflectCacherUtil;
import com.github.excel.util.StringUtil;
import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicLong;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:17 下午
 * @Description: Excel 通用服务
 */
@Slf4j
public class BaseExcelWriter {
	/**
	 * 生成Excel数据，list模型
	 */
	protected List<ExcelExportModel> exportModelList = new ArrayList<>();
	/**
	 * 生成Excel数据，单个bean模型
	 */
	protected List<ExcelExportModel> exportBeanList = new ArrayList<>();
	/**
	 * 自定义填充单元格List
	 */
	protected List<ExcelCustomColumnModel> customColumnModelList = new ArrayList<>();
	/**
	 * 自定义填充单元格List
	 */
	protected List<ExcelMergeCustomColumnModel> mergeCustomColumnModelList = new ArrayList<>();
	/**
	 * 自定义样式
	 */
	protected List<Class<? extends AbstractExcelStyle>> styleList = new ArrayList<>(ExcelConstant.TOW_INT);
	/**
	 * 样式缓存
	 */
	protected ThreadLocal<Map<String, CellStyle>> styleLocal = ThreadLocal.withInitial(() -> {
		return Maps.newHashMap();
	});

	/**
	 * 字体缓存
	 */
	protected ThreadLocal<Map<String, Font>> fontLocal = ThreadLocal.withInitial(() -> {
		return Maps.newHashMap();
	});
	/**
	 * 字段格式化器缓存
	 */
	protected Map<Class<? extends ExcelWriterDataFormat>, ExcelWriterDataFormat> dataFormatMap = new HashMap<>();
	/**
	 * Excel 名称
	 */
	protected String template;
	/**
	 * 自定义导出器
	 */
	protected ExcelCustomWriter customWrite;
	/**
	 * 默认格式化器
	 */
	protected ExcelWriterDataFormat dataFormat = new ExcelDefaultWriterDataFormat();
	/**
	 * class 信息缓存
	 */
	protected ReflectCacherUtil reflectCacherUtil = new ReflectCacherUtil();
	/**
	 * 是否启用流模式
	 */
	protected boolean streaming = false;
	/**
	 * 排除字段map
	 */
	protected Map<Class<? extends ExcelBaseModel>, Map<String, String>> excludeFieldMap = new HashMap<>();
	/**
	 * 校验或批注map
	 */
	protected Map<Class<? extends ExcelBaseModel>, Map<String, CommentModel>> commentMap = new HashMap<>();
	/**
	 * 设置选中sheet
	 */
	protected String selectSheet = null;
	/**
	 * 设置选中sheet
	 */
	protected Class<? extends ExcelBaseModel> listCla = null;

	/**
	 * 填充没有数据提示
	 */
	protected boolean noneDataTips = true;

	/**
	 * 自增map
	 */
	protected Map<Class<? extends ExcelBaseModel>, AtomicLong> incrementSeqMap = Maps.newConcurrentMap();

	/**
	 * 根据模板导出
	 *
	 * @param outputStream 导出流
	 */
	protected void writeToTemplate(OutputStream outputStream) {
		ExcelTemplateCacheModel excelTemplateCacheModel = ExcelBootLoader.getExcelTemplateFileCacheMapValue(template);
		Map<String, List<ExcelExpressionModel>> expressionMap = ExcelBootLoader.getExcelTemplateCacheMapValue(template);
		if (Objects.isNull(expressionMap) || Objects.isNull(excelTemplateCacheModel)) {
			throw new ExcelWriteException("excel template not found");
		}
		ThreadLocal<WorkbookCachePool.WorkbookCacheModel> workbookThreadLocal = excelTemplateCacheModel.getWorkbookThreadLocal();
		if (Objects.isNull(workbookThreadLocal)) {
			throw new ExcelWriteException("Failed to fetch workbook from cache");
		}
		WorkbookCachePool.WorkbookCacheModel workbookCacheModel = workbookThreadLocal.get();
		fontLocal.set(workbookCacheModel.getFontMap());
		styleLocal.set(workbookCacheModel.getStyleMap());

		try (Workbook workbook = workbookCacheModel.getWorkbook()) {
			CreationHelper creationHelper = workbook.getCreationHelper();
			initStyle(workbook);

			Map<Integer, List<ExcelTemplateTitleModel>> titleMap = ExcelBootLoader.getExcelTemplateTitleCacheMapValue(template);
			// 模板导出
			fillByTemplate(workbook, expressionMap, titleMap, creationHelper);

			if (Objects.nonNull(customWrite)) {
				customWrite.execute(workbook);
			}
			addNoResultData(workbook, creationHelper);
			selectSheet(workbook);
			workbook.write(outputStream);
		} catch (IOException e) {
			log.error("Export excel failed, cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		} finally {
			workbookThreadLocal.remove();
			styleLocal.remove();
			fontLocal.remove();
		}
	}


	/**
	 * 导出到文件
	 *
	 * @param outputStream 导出流
	 */
	protected void writeToNewFile(OutputStream outputStream, String excelName) {
		SXSSFWorkbook sxssfWorkbook = null;
		ThreadLocal<WorkbookCachePool.WorkbookCacheModel> workbookThreadLocal;
		if (excelName.endsWith(ExcelConstant.XLSX_STR)) {
			if (streaming) {
				workbookThreadLocal = WorkbookCachePool.getSxssfWorkbookThreadLocal();
			} else {
				workbookThreadLocal = WorkbookCachePool.getXssfWorkbookThreadLocal();
			}
		} else {
			workbookThreadLocal = WorkbookCachePool.getHssfWorkbookThreadLocal();
		}

		if (Objects.isNull(workbookThreadLocal)) {
			throw new ExcelWriteException("Failed to fetch workbook from cache");
		}
		WorkbookCachePool.WorkbookCacheModel cacheModel = workbookThreadLocal.get();
		styleLocal.set(cacheModel.getStyleMap());
		fontLocal.set(cacheModel.getFontMap());

		try (Workbook workbook = cacheModel.getWorkbook()) {

			if (workbook instanceof SXSSFWorkbook) {
				sxssfWorkbook = (SXSSFWorkbook) workbook;
			}
			// 初始化样式
			initStyle(workbook);
			CreationHelper createHelper = workbook.getCreationHelper();

			fillCustomColumn(workbook, createHelper);
			fillMergeCustomColumn(workbook, createHelper);
			for (ExcelExportModel exportModel : exportBeanList) {
				Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, exportModel.getSheetName());
				fillBean(workbook, createHelper, exportModel, sheet);
			}
			for (ExcelExportModel exportModel : exportModelList) {
				Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, exportModel.getSheetName());
				fillBeanList(workbook, createHelper, exportModel, sheet);
			}

			if (Objects.nonNull(customWrite)) {
				customWrite.execute(workbook);
			}
			addNoResultData(workbook, createHelper);
			selectSheet(workbook);
			workbook.write(outputStream);
		} catch (IllegalAccessException e) {
			log.error("Export excel failed, cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		} catch (IOException e) {
			log.error("Export excel failed, cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		} catch (InvocationTargetException e) {
			log.error("Export excel failed, cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		} finally {
			if (Objects.nonNull(sxssfWorkbook)) {
				sxssfWorkbook.dispose();
			}
			styleLocal.remove();
			fontLocal.remove();
		}
	}

	/**
	 * 设置sheet为选中状态
	 *
	 * @param workbook
	 */
	protected void selectSheet(Workbook workbook) {
		if (StringUtil.notEmpty(selectSheet)) {
			Sheet sheet = workbook.getSheet(selectSheet);
			if (Objects.nonNull(sheet)) {
				sheet.setSelected(true);
			}
		}
	}

	/**
	 * 执行模板填充
	 *
	 * @param workbook
	 * @param expressionMap
	 * @param titleMap
	 * @param creationHelper
	 */
	protected void fillByTemplate(Workbook workbook, Map<String, List<ExcelExpressionModel>> expressionMap, Map<Integer, List<ExcelTemplateTitleModel>> titleMap, CreationHelper creationHelper) {
		try {
			// 填充单列
			fillCustomColumn(workbook, creationHelper);
			fillMergeCustomColumn(workbook, creationHelper);
			for (ExcelExportModel exportModel : exportBeanList) {
				// 填充单个bean
				if (exportModel.getFillTemplate()) {
					fillBeanByTemplate(workbook, expressionMap, creationHelper, exportModel);
				} else {
					Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, exportModel.getSheetName());
					fillBean(workbook, creationHelper, exportModel, sheet);
				}
			}
			for (ExcelExportModel exportModel : exportModelList) {
				// 填充列表
				if (exportModel.getFillTemplate()) {
					fillBeanListByTemplate(workbook, titleMap, creationHelper, exportModel);
				} else {
					Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, exportModel.getSheetName());
					fillBeanList(workbook, creationHelper, exportModel, sheet);
				}

			}
		} catch (Exception e) {
			log.error("Export excel by template failed,cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException("Export excel by template failed");
		}
	}

	/**
	 * 根据填充列表
	 *
	 * @param workbook
	 * @param titleMap
	 * @param creationHelper
	 * @param exportModel
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	protected void fillBeanListByTemplate(Workbook workbook, Map<Integer, List<ExcelTemplateTitleModel>> titleMap, CreationHelper creationHelper, ExcelExportModel exportModel) throws IllegalAccessException, InvocationTargetException {
		// 确定映射关系
		ExcelCacheModel cacheModel = exportModel.getCacheModel();
		List<ExcelListTemplateDto> templateDtoList = Lists.newArrayList();
		Map<String, String> excludeMap = excludeFieldMap.get(exportModel.getExcelModelClass());
		for (Map.Entry<Integer, List<ExcelTemplateTitleModel>> titleRowMap : titleMap.entrySet()) {
			// 循环每一行标题
			for (ExcelTemplateTitleModel titleModel : titleRowMap.getValue()) {
				ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel = cacheModel.getFieldModelMap().get(titleModel.getTitle());
				if (Objects.nonNull(cacheFieldModel)) {
					if (Objects.nonNull(excludeMap) && Objects.nonNull(excludeMap.get(cacheFieldModel.getFieldName()))) {
						continue;
					}
					ExcelListTemplateDto listTemplateDto = ExcelListTemplateDto.builder().rowIndex(titleModel.getRowIndex()).cacheFieldModel(cacheFieldModel).exportCell(cacheFieldModel.getExportCell()).colIndex(titleModel.getColIndex()).getMethod(cacheFieldModel.getGetMethod()).sheetName(titleModel.getSheetName()).build();
					templateDtoList.add(listTemplateDto);
				}
			}
		}
		if (templateDtoList.size() == ExcelConstant.ZERO_SHORT) {
			return;
		}
		// 填充列表bean 以list为准
		int rowIndex = templateDtoList.get(ExcelConstant.ZERO_SHORT).getRowIndex() + ExcelConstant.ONE_INT;
		Sheet sheet = workbook.getSheet(templateDtoList.get(ExcelConstant.ZERO_SHORT).getSheetName());
		// 移动行
		if (cacheModel.getExcelExport().fillType() == ExcelExportListFillTypeEnum.SHIFT) {
			ExcelHelper.shiftRows(sheet, rowIndex, exportModel.getDataModelList().size());
		}
		int i = ExcelConstant.ZERO_SHORT;
		for (ExcelBaseModel baseModel : exportModel.getDataModelList()) {
			i++;
			Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
			for (ExcelListTemplateDto templateDto : templateDtoList) {
				String styleName = i % ExcelConstant.TOW_INT == ExcelConstant.ZERO_SHORT ? templateDto.getCacheFieldModel().getEvenRowStyleName() : templateDto.getCacheFieldModel().getContentStyleName();
				Object value = templateDto.getGetMethod().invoke(baseModel);
				Cell cell = ExcelHelper.getCellOrCreate(row, templateDto.getColIndex());
				ExcelHelper.setRowHeight(row, templateDto.getCacheFieldModel().getContentRowHeight());
				ExcelHelper.setColWidth(sheet, cell.getColumnIndex(), templateDto.getCacheFieldModel().getColWidth());

				ExcelWriterDataFormat formatter = getFormatter(templateDto.getExportCell().formatter());
				value = formatValue(value, templateDto.getExportCell().formatPattern(), formatter);
				if (Objects.nonNull(value) && value instanceof byte[]) {
					createPicture(workbook, sheet, cell, (byte[]) value, styleName);
				} else {
					setCellValueAndStyle(cell, value, styleName, templateDto.getExportCell().linkName(), creationHelper);
				}
			}
			rowIndex++;
		}
	}

	/**
	 * 根据模板填充单个bean
	 *
	 * @param workbook
	 * @param expressionMap
	 * @param creationHelper
	 * @param exportModel
	 * @throws Exception
	 */
	protected void fillBeanByTemplate(Workbook workbook, Map<String, List<ExcelExpressionModel>> expressionMap, CreationHelper creationHelper, ExcelExportModel exportModel) throws Exception {
		ExcelExport excelExport = exportModel.getCacheModel().getExcelExport();
		List<ExcelExpressionModel> expressionModelList = expressionMap.get(excelExport.nameSpace());

		if (Objects.nonNull(expressionModelList)) {
			Map<String, String> excludeMap = excludeFieldMap.get(exportModel.getExcelModelClass());
			// 替换表达式
			for (ExcelExpressionModel expressionModel : expressionModelList) {
				List<ExcelCacheModel.ExcelCacheFieldModel> fieldModelList = exportModel.getCacheModel().getFieldModelList();

				for (ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel : fieldModelList) {
					// 单个属性
					if (cacheFieldModel.getFieldName().equals(expressionModel.getFieldName()[ExcelConstant.ZERO_SHORT])) {
						// 执行填充
						ExcelExportCell exportCell = cacheFieldModel.getExportCell();
						String styleName = StringUtil.notEmpty(exportCell.contentStyleName()) ? exportCell.contentStyleName() : excelExport.contentStyleName();

						Sheet sheet = workbook.getSheet(expressionModel.getSheetName());
						Row row = ExcelHelper.getRowOrCreate(sheet, expressionModel.getRowIndex());
						Cell cell = ExcelHelper.getCellOrCreate(row, expressionModel.getColIndex());
						ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
						ExcelHelper.setColWidth(sheet, cell.getColumnIndex(), cacheFieldModel.getColWidth());

						Object value;
						if (Objects.nonNull(excludeMap) && Objects.nonNull(excludeMap.get(expressionModel.getFieldName()[ExcelConstant.ZERO_SHORT]))) {
							value = null;
						} else {
							value = cacheFieldModel.getGetMethod().invoke(exportModel.getDataModel());
						}
						if (expressionModel.getFieldName().length > ExcelConstant.ONE_INT) {
							for (int i = ExcelConstant.ONE_INT; i < expressionModel.getFieldName().length; i++) {
								value = reflectCacherUtil.getObjectThenCache(value, expressionModel.getFieldName()[i]);
							}
						}
						String expression = ExcelConstant.EXPRESSION_PREFIX + expressionModel.getExpressionContent() + ExcelConstant.EXPRESSION_SUFFIX;
						if (Objects.nonNull(value) && value instanceof byte[]) {
							createPicture(workbook, sheet, cell, (byte[]) value, styleName);
							replaceExpression(cell, ExcelConstant.NULL_STR, expression);
						} else {
							ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());
							value = formatValue(value, exportCell.formatPattern(), formatter);
							setCellValueAndStyle(cell, value, styleName, exportCell.linkName(), creationHelper, expression);
						}
						break;
					}
				}

			}
		}
	}


	/**
	 * 初始化样式
	 *
	 * @param workbook workBook
	 */
	protected void initStyle(Workbook workbook) {
		try {
			for (Class<? extends AbstractExcelStyle> styleClass : styleList) {
				initStyle(workbook, styleClass);
			}
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException("Init style error");
		}
	}

	/**
	 * 初始化样式
	 *
	 * @param workbook   workBook
	 * @param styleClass styleClass
	 */
	protected void initStyle(Workbook workbook, Class<? extends AbstractExcelStyle> styleClass) {
		try {

			Constructor<? extends AbstractExcelStyle> constructor = styleClass.getConstructor(styleClass.getConstructors()[ExcelConstant.ZERO_SHORT].getParameterTypes());
			AbstractExcelStyle excelStyle = constructor.newInstance(workbook, styleLocal.get(), fontLocal.get());
			excelStyle.addNewFont();
			excelStyle.addNewStyle();
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException("Init style error");
		}
	}

	/**
	 * 获取格式化器
	 *
	 * @param formatClass 格式化器class
	 */
	protected ExcelWriterDataFormat getFormatter(Class<? extends ExcelWriterDataFormat> formatClass) {
		if (formatClass == ExcelDefaultWriterDataFormat.class) {
			return null;
		}
		ExcelWriterDataFormat format = dataFormatMap.get(formatClass);
		if (null == format) {
			try {
				format = formatClass.newInstance();
			} catch (Exception e) {
				log.error("Error by create formatter , cause:{}", Throwables.getStackTraceAsString(e));
			}
			dataFormatMap.put(formatClass, format);
		}
		return format;
	}

	/**
	 * 填充自定义列
	 *
	 * @param workbook
	 * @param createHelper
	 */
	protected void fillCustomColumn(Workbook workbook, CreationHelper createHelper) {
		for (ExcelCustomColumnModel columnModel : customColumnModelList) {
			fillColumn(workbook, createHelper, columnModel, columnModel.getRowIndex(), columnModel.getColIndex());
		}
	}

	/**
	 * 填充column
	 *
	 * @param workbook
	 * @param createHelper
	 * @param columnModel
	 * @param rowIndex
	 * @param colIndex
	 */
	protected void fillColumn(Workbook workbook, CreationHelper createHelper, BaseExcelColumnModel columnModel, int rowIndex, int colIndex) {
		Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, columnModel.getSheetName());
		Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setColWidth(sheet, colIndex, columnModel.getColWidth());
		ExcelHelper.setRowHeight(row, columnModel.getRowHeight());
		ExcelWriterDataFormat formatter = getFormatter(columnModel.getDataFormat());
		Object value = formatValue(columnModel.getValue(), columnModel.getFormatPattern(), formatter);

		if (columnModel.getFillType() == ExcelExportColumnFillTypeEnum.APPEND) {
			cell.setCellType(CellType.STRING);
			String cellValue = cell.getStringCellValue();
			if (StringUtil.notEmpty(cellValue)) {
				value = cellValue + value;
			}
		}
		setCellValueAndStyle(cell, value, columnModel.getStyleName(), null, createHelper);
	}

	/**
	 * 填充合并列
	 *
	 * @param workbook
	 * @param createHelper
	 */
	protected void fillMergeCustomColumn(Workbook workbook, CreationHelper createHelper) {
		for (ExcelMergeCustomColumnModel columnModel : mergeCustomColumnModelList) {
			Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, columnModel.getSheetName());
			sheet.addMergedRegion(new CellRangeAddress(columnModel.getFirstRow(), columnModel.getLastRow(), columnModel.getFirstColumn(), columnModel.getLastColumn()));
			fillColumn(workbook, createHelper, columnModel, columnModel.getFirstRow(), columnModel.getFirstColumn());
		}
	}

	/**
	 * 填充单个bean
	 *
	 * @param workbook     workbook
	 * @param createHelper 创建对象的
	 * @param exportModel  数据组装实体
	 * @param sheet        sheet
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	protected void fillBean(Workbook workbook, CreationHelper createHelper, ExcelExportModel exportModel, Sheet sheet) throws IllegalAccessException, InvocationTargetException {
		ExcelCacheModel cacheModel = exportModel.getCacheModel();
		ExcelExport excelExport = cacheModel.getExcelExport();
		int colIndex = excelExport.colIndex(), rowIndex = excelExport.rowIndex(), initColIndex = excelExport.colIndex();
		if (null != exportModel.getRowIndex()) {
			rowIndex = exportModel.getRowIndex();
		}

		if (null != exportModel.getColIndex()) {
			colIndex = exportModel.getColIndex();
			initColIndex = exportModel.getColIndex();
		}

		doFillBean(workbook, createHelper, exportModel, sheet, cacheModel, excelExport, colIndex, rowIndex, initColIndex);
	}

	/**
	 * 执行填充单个对象
	 *
	 * @param workbook
	 * @param createHelper
	 * @param exportModel
	 * @param sheet
	 * @param cacheModel
	 * @param excelExport
	 * @param colIndex
	 * @param rowIndex
	 * @param initColIndex
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	protected void doFillBean(Workbook workbook, CreationHelper createHelper, ExcelExportModel exportModel, Sheet sheet, ExcelCacheModel cacheModel, ExcelExport excelExport, int colIndex, int rowIndex, final int initColIndex) throws IllegalAccessException, InvocationTargetException {
		boolean hasCycle = false;
		ExcelExportFillStyleEnum lasFileStyle = null;
		List<Integer> rowIndexList = Lists.newArrayList();
		Map<String, String> excludeMap = excludeFieldMap.get(exportModel.getExcelModelClass());
		for (ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel : cacheModel.getFieldModelList()) {
			boolean isMap = cacheFieldModel.isMap();
			ExcelExportCell exportCell = cacheFieldModel.getExportCell();
			if (Objects.nonNull(excludeMap) && Objects.nonNull(excludeMap.get(cacheFieldModel.getFieldName()))) {
				continue;
			}
			ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());

			Object value = cacheFieldModel.getGetMethod().invoke(exportModel.getDataModel());
			value = formatValue(value, exportCell.formatPattern(), formatter);

			// 填充指定的列
			if (exportCell.rowIndex() != ExcelConstant.MINUS_ONE_SHORT && exportCell.colIndex() != ExcelConstant.MINUS_ONE_SHORT) {
				Row row = ExcelHelper.getRowOrCreate(sheet, exportCell.rowIndex());
				ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
				fillContent(workbook, sheet, exportCell.colIndex(), rowIndex, exportCell, value, row, createHelper, excelExport, isMap, cacheFieldModel);
			} else {
				// 自动填充
				if (hasCycle) {
					if (ExcelExportFillStyleEnum.HORIZONTAL == exportCell.fillStyle()) {
						colIndex++;
						if (exportCell.verticalNewLine() && ExcelExportFillStyleEnum.VERTICAL == lasFileStyle) {
							rowIndex++;
							colIndex = initColIndex;
						}
					} else if (ExcelExportFillStyleEnum.VERTICAL == exportCell.fillStyle()) {
						if (rowIndexList.size() > ExcelConstant.ZERO_SHORT) {
							rowIndex = rowIndexList.stream().max(Comparator.comparing(Integer::intValue)).get();
							rowIndexList.clear();
						}
						rowIndex++;
						colIndex = initColIndex;
					}
				}
				Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
				ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
				hasCycle = true;
				lasFileStyle = exportCell.fillStyle();

				ColPointDto pointDto = fillContent(workbook, sheet, colIndex, rowIndex, exportCell, value, row, createHelper, excelExport, isMap, cacheFieldModel);
				colIndex = pointDto.colIndex;
				if (ExcelExportFillStyleEnum.HORIZONTAL == exportCell.fillStyle()) {
					rowIndexList.add(pointDto.rowIndex);
				} else if (ExcelExportFillStyleEnum.VERTICAL == exportCell.fillStyle()) {
					rowIndex = pointDto.rowIndex;
				}
			}
		}
	}

	/**
	 * 填充列表
	 *
	 * @param workbook     workbook
	 * @param createHelper 创建对象的
	 * @param exportModel  数据组装实体
	 * @param sheet        sheet
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	protected void fillBeanList(Workbook workbook, CreationHelper createHelper, ExcelExportModel exportModel, Sheet sheet) {
		if (exportModel.getDataModelList().size() > ExcelConstant.ZERO_SHORT) {
			ExcelCacheModel cacheModel = exportModel.getCacheModel();
			int rowIndex = cacheModel.getExcelExport().rowIndex(), colIndex = cacheModel.getExcelExport().colIndex(), initColIndex = cacheModel.getExcelExport().colIndex(), initRowIndex = cacheModel.getExcelExport().rowIndex();
			ExcelExportFillStyleEnum fillStyleEnum = exportModel.getCacheModel().getExcelExport().fillStyle();
			if (null != exportModel.getRowIndex()) {
				rowIndex = exportModel.getRowIndex();
				initRowIndex = exportModel.getRowIndex();
			}

			if (null != exportModel.getColIndex()) {
				colIndex = exportModel.getColIndex();
				initColIndex = exportModel.getColIndex();
			}
			// 移动行
			if (cacheModel.getExcelExport().fillType() == ExcelExportListFillTypeEnum.SHIFT) {
				ExcelHelper.shiftRows(sheet, rowIndex, cacheModel.getExcelExport().fillTitle() ? exportModel.getDataModelList().size() + ExcelConstant.ONE_INT : exportModel.getDataModelList().size());
			}
			List<ExcelCacheModel.ExcelCacheFieldModel> fillFieldModelList = new ArrayList<>();
			Map<String, String> excludeMap = excludeFieldMap.get(exportModel.getExcelModelClass());
			Map<String, Integer> commentPointIndexMap = new HashMap<>();
			int i = colIndex;
			for (ExcelCacheModel.ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
				if (Objects.nonNull(excludeMap) && Objects.nonNull(excludeMap.get(fieldModel.getFieldName()))) {
					continue;
				}
				commentPointIndexMap.put(fieldModel.getFieldName(), i);
				fillFieldModelList.add(fieldModel);
				i++;
			}
			//填充标题
			ColPointDto colPointDto = fillListTitle(cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, initColIndex, initRowIndex, exportModel, fillFieldModelList);
			rowIndex = colPointDto.getRowIndex();
			colIndex = colPointDto.getColIndex();

			// 填充内容
			ColPointDto contentLastPoint = fillListContent(workbook, exportModel, cacheModel, sheet, createHelper, fillFieldModelList, rowIndex, colIndex, fillStyleEnum, initColIndex, initRowIndex);
			fillListValidationOrComment(workbook, sheet, rowIndex, contentLastPoint.getRowIndex(), commentPointIndexMap, commentMap.get(exportModel.getExcelModelClass()), exportModel);
		}
	}

	/**
	 * 添加校验或备注
	 *
	 * @param sheet
	 * @param startRowIndex
	 * @param endRowIndex
	 * @param commentColIndexMap
	 * @param commentModelMap
	 */
	protected void fillListValidationOrComment(Workbook workbook, Sheet sheet, int startRowIndex, int endRowIndex, Map<String, Integer> commentColIndexMap, Map<String, CommentModel> commentModelMap, ExcelExportModel exportModel) {
		if (Objects.isNull(commentModelMap)) {
			return;
		}
		if (exportModel.getCacheModel().getExcelExport().fillStyle() != ExcelExportFillStyleEnum.VERTICAL) {
			return;
		}
		for (Map.Entry<String, CommentModel> entry : commentModelMap.entrySet()) {
			Integer colIndex = commentColIndexMap.get(entry.getKey());
			if (Objects.isNull(colIndex)) {
				continue;
			}
			CommentModel commentModel = entry.getValue();
			if (commentModel instanceof ComboBoxModel) {
				ExcelHelper.addDropDownValidation(sheet, new CellRangeAddressList(startRowIndex, endRowIndex - ExcelConstant.ONE_INT, colIndex, colIndex), ((ComboBoxModel) commentModel).getOptions());
			} else if (commentModel instanceof NumberScopeModel) {
				NumberScopeModel numberScopeModel = (NumberScopeModel) commentModel;
				ExcelHelper.addRangeValidation(sheet, new CellRangeAddressList(startRowIndex, endRowIndex - ExcelConstant.ONE_INT, colIndex, colIndex), numberScopeModel.getStart(), numberScopeModel.getEnd());
			}
			if (exportModel.getCacheModel().getExcelExport().fillTitle() && StringUtil.notEmpty(commentModel.getCommentText())) {
				ExcelHelper.createComment(workbook, sheet, startRowIndex - ExcelConstant.ONE_INT, colIndex, ExcelConstant.NULL_STR, commentModel.getCommentText(), fontLocal.get().get(commentModel.getCommentFontName()));
			}
		}
	}

	protected ColPointDto fillListMapTitle(ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelExportFillStyleEnum fillStyleEnum, ExcelCacheModel.ExcelCacheFieldModel fieldModel, ExcelExportModel exportModel) {
		try {
			ExcelBaseModel excelBaseModel = exportModel.getDataModelList().get(ExcelConstant.ZERO_SHORT);
			ExcelExportCell exportCell = fieldModel.getExportCell();
			Object value = fieldModel.getGetMethod().invoke(excelBaseModel);
			if (value instanceof Map) {
				Iterator keyIterator = ((Map) value).keySet().iterator();
				while (keyIterator.hasNext()) {
					String title = "";
					if (StringUtil.notEmpty(exportCell.titleName())) {
						title = title + exportCell.titleName();
					}
					String key = keyIterator.next().toString();
					if (Strings.isNullOrEmpty(title)) {
						title = key;
					} else {
						title = title + ExcelConstant.SEPARATOR + key;
					}

					String styleName = StringUtil.notEmpty(exportCell.titleStyleName()) ? exportCell.titleStyleName() : cacheModel.getExcelExport().titleStyleName();
					Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
					Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
					ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
					ExcelHelper.setRowHeight(row, fieldModel.getContentRowHeight());
					setCellValueAndStyle(cell, title, styleName, null, createHelper);

					if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
						colIndex++;
					} else {
						rowIndex++;
					}
				}
			}
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			//invoke方法执行出错
			throw new ExcelWriteException("method.invoke.fail");
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	protected ColPointDto fillListBeanTitle(ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelExportFillStyleEnum fillStyleEnum, ExcelCacheModel.ExcelCacheFieldModel fieldModel) {
		ExcelExportCell exportCell = fieldModel.getExportCell();
		String titleName ;
		if (Objects.nonNull(exportCell)) {
			titleName = exportCell.titleName();
		}else{
			titleName = cacheModel.getExcelExport().incrementSequenceTitle();
		}

		Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
		ExcelHelper.setRowHeight(row, fieldModel.getTitleRowHeight());
		setCellValueAndStyle(cell, titleName, fieldModel.getTitleStyleName(), null, createHelper);
		// merge column
		if (cacheModel.getExcelExport().mergeTitleRowNum() > ExcelConstant.ZERO_SHORT || cacheModel.getExcelExport().mergeTitleColNum() > ExcelConstant.ZERO_SHORT) {
			int mergeRowEndIndex = rowIndex + cacheModel.getExcelExport().mergeTitleRowNum();
			int mergeColEndIndex = colIndex + cacheModel.getExcelExport().mergeTitleColNum();

			sheet.addMergedRegion(new CellRangeAddress(rowIndex, mergeRowEndIndex, colIndex, mergeColEndIndex));

			if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
				colIndex = mergeColEndIndex;
			} else {
				rowIndex = mergeRowEndIndex;
			}

		}
		if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
			colIndex++;
		} else {
			rowIndex++;
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充列表标题
	 *
	 * @param cacheModel
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param initColIndex
	 * @param initRowIndex
	 * @param fillFieldModelList
	 * @return
	 */
	protected ColPointDto fillListTitle(ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelExportFillStyleEnum fillStyleEnum, final int initColIndex, final int initRowIndex, ExcelExportModel exportModel, final List<ExcelCacheModel.ExcelCacheFieldModel> fillFieldModelList) {
		if (cacheModel.getExcelExport().fillTitle()) {
			for (ExcelCacheModel.ExcelCacheFieldModel fieldModel : fillFieldModelList) {
				if (fieldModel.isMap()) {
					ColPointDto colPointDto = fillListMapTitle(cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, fieldModel, exportModel);
					rowIndex = colPointDto.rowIndex;
					colIndex = colPointDto.colIndex;
				} else {
					ColPointDto colPointDto = fillListBeanTitle(cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, fieldModel);
					rowIndex = colPointDto.rowIndex;
					colIndex = colPointDto.colIndex;
				}
			}
			if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
				rowIndex += cacheModel.getExcelExport().mergeTitleRowNum() + ExcelConstant.ONE_INT;
			} else {
				colIndex += cacheModel.getExcelExport().mergeTitleColNum() + ExcelConstant.ONE_INT;
			}
			if (cacheModel.getExcelExport().freezeTitle()) {
				sheet.createFreezePane(ExcelConstant.ZERO_SHORT, cacheModel.getExcelExport().rowIndex() + ExcelConstant.ONE_INT, ExcelConstant.ZERO_SHORT, cacheModel.getExcelExport().rowIndex() + ExcelConstant.ONE_INT);
			}
		}

		if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
			colIndex = initColIndex;
		} else {
			rowIndex = initRowIndex;
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充列表list字段
	 *
	 * @param workbook
	 * @param cacheFieldModel
	 * @param cacheModel
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param value
	 * @param styleName
	 * @return
	 */
	protected ColPointDto fillListField(Workbook workbook, ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel, ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelExportFillStyleEnum fillStyleEnum, Object value, String styleName) {
		ExcelExportCell exportCell = cacheFieldModel.getExportCell();
		String linkName = ExcelConstant.NULL_STR;
		if (Objects.nonNull(exportCell)) {
			linkName = exportCell.linkName();
		}
		Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
		ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
		if (Objects.nonNull(value) && value instanceof byte[]) {
			createPicture(workbook, sheet, cell, (byte[]) value, styleName);
		} else {
			setCellValueAndStyle(cell, value, styleName, linkName, createHelper);
		}
		// merge column   不判断标题是否合并
		if (canMergeColumn(cacheModel.getExcelExport(), fillStyleEnum)) {
			int mergeContentRowNum = cacheModel.getExcelExport().mergeContentRowNum(), mergeContentColNum = cacheModel.getExcelExport().mergeContentColNum();

			if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
				mergeContentColNum = cacheModel.getExcelExport().mergeTitleColNum();
			} else if (ExcelExportFillStyleEnum.HORIZONTAL == fillStyleEnum) {
				mergeContentRowNum = cacheModel.getExcelExport().mergeTitleRowNum();
			}
			int mergeRowEndIndex = rowIndex + mergeContentRowNum;
			int mergeColEndIndex = colIndex + mergeContentColNum;
			sheet.addMergedRegion(new CellRangeAddress(rowIndex, mergeRowEndIndex, colIndex, mergeColEndIndex));
			if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
				colIndex = mergeColEndIndex;
			} else {
				rowIndex = mergeRowEndIndex;
			}
		}
		if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
			colIndex++;
		} else {
			rowIndex++;
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 是否可以合并list列表单元格
	 *
	 * @param excelExport
	 * @param fillStyleEnum
	 * @return
	 */
	protected boolean canMergeColumn(ExcelExport excelExport, ExcelExportFillStyleEnum fillStyleEnum) {
		if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
			if (excelExport.mergeTitleColNum() > ExcelConstant.ZERO_SHORT || excelExport.mergeContentColNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			} else if (excelExport.mergeContentRowNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			}
		} else if (ExcelExportFillStyleEnum.HORIZONTAL == fillStyleEnum) {
			if (excelExport.mergeTitleRowNum() > ExcelConstant.ZERO_SHORT || excelExport.mergeContentRowNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			} else if (excelExport.mergeContentColNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			}
		}
		return false;
	}

	/**
	 * 填充列表map字段
	 *
	 * @param workbook
	 * @param exportCell
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param value
	 * @param styleName
	 * @param cacheFieldModel
	 * @return
	 */
	protected ColPointDto fillListMapField(Workbook workbook, ExcelExportCell exportCell, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelExportFillStyleEnum fillStyleEnum, Object value, String styleName, ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel) {
		if (value instanceof Map) {
			Iterator keyIterator = ((Map) value).keySet().iterator();
			while (keyIterator.hasNext()) {
				Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
				Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
				ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				Object cellValue = ((Map) value).get(keyIterator.next());
				if (Objects.nonNull(cellValue) && cellValue instanceof byte[]) {
					createPicture(workbook, sheet, cell, (byte[]) cellValue, styleName);
				} else {
					setCellValueAndStyle(cell, cellValue, styleName, exportCell.linkName(), createHelper);
				}
				if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
					colIndex++;
				} else {
					rowIndex++;
				}
			}
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充列表内容
	 *
	 * @param workbook
	 * @param exportModel
	 * @param cacheModel
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param initColIndex
	 * @param initRowIndex
	 * @return
	 */
	protected ColPointDto fillListContent(Workbook workbook, ExcelExportModel exportModel, ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, List<ExcelCacheModel.ExcelCacheFieldModel> fillFieldModelList, int rowIndex, int colIndex, ExcelExportFillStyleEnum fillStyleEnum, final int initColIndex, final int initRowIndex) {
		int contentRow = ExcelConstant.ZERO_SHORT;
		// 填充内容
		for (ExcelBaseModel model : exportModel.getDataModelList()) {
			contentRow++;
			try {
				for (ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel : fillFieldModelList) {
					Object value ;
					if (Objects.nonNull(cacheFieldModel.getGetMethod())) {
						value = cacheFieldModel.getGetMethod().invoke(model);
					}else{
						AtomicLong incrementSeq = incrementSeqMap.get(exportModel.getExcelModelClass());
						if (incrementSeq == null) {
							incrementSeq = new AtomicLong(ExcelConstant.ZERO_SHORT);
							incrementSeqMap.put(exportModel.getExcelModelClass(), incrementSeq);
						}
						value = incrementSeq.incrementAndGet();
					}
					ExcelExportCell exportCell = cacheFieldModel.getExportCell();
					String styleName = contentRow % ExcelConstant.TOW_INT == ExcelConstant.ZERO_SHORT ? cacheFieldModel.getEvenRowStyleName() : cacheFieldModel.getContentStyleName();

					if(Objects.nonNull(exportCell)) {
						ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());
						value = formatValue(value, exportCell.formatPattern(), formatter);
					}

					if (cacheFieldModel.isMap()) {
						ColPointDto colPointDto = fillListMapField(workbook, exportCell, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, value, styleName, cacheFieldModel);
						rowIndex = colPointDto.rowIndex;
						colIndex = colPointDto.colIndex;
					} else {
						ColPointDto colPointDto = fillListField(workbook, cacheFieldModel, cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, value, styleName);
						rowIndex = colPointDto.rowIndex;
						colIndex = colPointDto.colIndex;
					}

				}
				if (ExcelExportFillStyleEnum.VERTICAL == fillStyleEnum) {
					rowIndex += cacheModel.getExcelExport().mergeContentRowNum() + ExcelConstant.ONE_INT;
					colIndex = initColIndex;
				} else {
					colIndex += cacheModel.getExcelExport().mergeContentColNum() + ExcelConstant.ONE_INT;
					rowIndex = initRowIndex;
				}
			} catch (Exception e) {
				log.info(Throwables.getStackTraceAsString(e));
				throw new ExcelWriteException(e.getMessage());
			}
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充内容
	 *
	 * @param wb              Workbook
	 * @param sheet           sheet 名称
	 * @param colIndex        列宽
	 * @param exportCell      cell
	 * @param value           value
	 * @param row             row
	 * @param createHelper    excelHelper
	 * @param excelExport
	 * @param cacheFieldModel
	 * @return int
	 */
	protected ColPointDto fillContent(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelExportCell exportCell, Object value, Row row, CreationHelper createHelper, ExcelExport excelExport, boolean isMap, ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel) {

		ExcelExportCellTitleModelEnum titleModelEnum = excelExport.titleModel();
		if (titleModelEnum == ExcelExportCellTitleModelEnum.DEFAULT) {
			titleModelEnum = exportCell.titleModel();
		} else {
			titleModelEnum = exportCell.titleModel() == ExcelExportCellTitleModelEnum.DEFAULT ? titleModelEnum : exportCell.titleModel();
		}

		if (titleModelEnum == ExcelExportCellTitleModelEnum.STAND_ALONE || titleModelEnum == ExcelExportCellTitleModelEnum.DEFAULT) {
			return fillStandAloneContent(wb, sheet, colIndex, rowIndex, exportCell, value, row, createHelper, excelExport, isMap, cacheFieldModel);
		} else if (titleModelEnum == ExcelExportCellTitleModelEnum.WITH_VALUE) {
			return fillWithValueContent(wb, sheet, colIndex, rowIndex, exportCell, value, row, createHelper, excelExport, cacheFieldModel);
		}
		return null;
	}

	/**
	 * 填充合并单元格
	 *
	 * @param exportCell
	 * @param rowIndex
	 * @param colIndex
	 * @param sheet
	 * @return
	 */
	protected ColPointDto cellMerged(ExcelExportCell exportCell, int rowIndex, int colIndex, Sheet sheet) {
		int mergeContentRowNum = exportCell.mergeRowNum(), mergeContentColNum = exportCell.mergeContentColNum();

		if (mergeContentRowNum > ExcelConstant.ZERO_SHORT || mergeContentColNum > ExcelConstant.ZERO_SHORT) {
			int mergeRowEndIndex = rowIndex + mergeContentRowNum;
			int mergeColEndIndex = colIndex + mergeContentColNum;
			sheet.addMergedRegion(new CellRangeAddress(rowIndex, mergeRowEndIndex, colIndex, mergeColEndIndex));
			colIndex = mergeColEndIndex;
			rowIndex = mergeRowEndIndex;
		}
		return ColPointDto.builder().rowIndex(rowIndex).colIndex(colIndex).build();
	}

	/**
	 * 填充合并单元格
	 *
	 * @param exportCell
	 * @param rowIndex
	 * @param colIndex
	 * @param sheet
	 * @return
	 */
	protected ColPointDto cellMergedTitle(ExcelExportCell exportCell, int rowIndex, int colIndex, Sheet sheet) {
		int mergeContentRowNum = exportCell.mergeRowNum(), mergeContentColNum = exportCell.mergeTitleColNum();

		if (mergeContentRowNum > ExcelConstant.ZERO_SHORT || mergeContentColNum > ExcelConstant.ZERO_SHORT) {
			int mergeRowEndIndex = rowIndex + mergeContentRowNum;
			int mergeColEndIndex = colIndex + mergeContentColNum;
			sheet.addMergedRegion(new CellRangeAddress(rowIndex, mergeRowEndIndex, colIndex, mergeColEndIndex));
			colIndex = mergeColEndIndex;
		}
		return ColPointDto.builder().rowIndex(rowIndex).colIndex(colIndex).build();
	}

	/**
	 * 填充标题和内容一起的内容
	 *
	 * @param wb
	 * @param sheet
	 * @param colIndex
	 * @param exportCell
	 * @param value
	 * @param row
	 * @param createHelper
	 * @param excelExport
	 * @param cacheFieldModel
	 * @return
	 */
	protected ColPointDto fillWithValueContent(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelExportCell exportCell, Object value, Row row, CreationHelper createHelper, ExcelExport excelExport, ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel) {
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
		if (Objects.nonNull(value) && value instanceof byte[]) {
			createPicture(wb, sheet, cell, (byte[]) value, cacheFieldModel.getTitleStyleName());
			return cellMerged(exportCell, rowIndex, colIndex, sheet);
		}

		if (StringUtil.notEmpty(exportCell.titleName())) {
			String content = exportCell.titleName() + exportCell.separator() + Optional.ofNullable(value).orElse(ExcelConstant.NULL_STR);
			int titleLen = exportCell.titleName().length();
			ExcelRichTextModel richTextModelTitle = null, richTextModelContent = null;
			CellStyle titleStyle = styleLocal.get().get(cacheFieldModel.getTitleStyleName());
			if (titleStyle != null) {
				Font titleFont = wb.getFontAt(titleStyle.getFontIndex());
				richTextModelTitle = new ExcelRichTextModel();
				richTextModelTitle.setFont(titleFont);
				richTextModelTitle.setStartIndex(ExcelConstant.ZERO_SHORT);
				richTextModelTitle.setEndIndex(titleLen);
			}

			CellStyle contentStyle = styleLocal.get().get(cacheFieldModel.getContentStyleName());
			if (contentStyle != null) {
				Font contentFont = wb.getFontAt(contentStyle.getFontIndex());
				richTextModelContent = new ExcelRichTextModel();
				richTextModelContent.setFont(contentFont);
				richTextModelContent.setStartIndex(titleLen);
				richTextModelContent.setEndIndex(content.length());
			}

			if (richTextModelContent == null && richTextModelTitle == null) {
				setCellValueAndStyle(cell, content, cacheFieldModel.getTitleStyleName(), null, createHelper);
			} else {
				RichTextString richText = ExcelHelper.createRichText(createHelper, content, richTextModelTitle, richTextModelContent);
				setCellValueAndStyle(cell, richText, cacheFieldModel.getTitleStyleName(), null, createHelper);
			}
		} else {
			setCellValueAndStyle(cell, value, cacheFieldModel.getTitleStyleName(), exportCell.linkName(), createHelper);
		}
		// 创建批注
		createComment(wb, sheet, colIndex, rowIndex, exportCell, value);
		return cellMerged(exportCell, rowIndex, colIndex, sheet);
	}

	/**
	 * 填充标题和内容单独占一列的内容
	 *
	 * @param wb
	 * @param sheet
	 * @param colIndex
	 * @param exportCell
	 * @param value
	 * @param row
	 * @param createHelper
	 * @param excelExport
	 * @param isMap
	 * @param cacheFieldModel
	 * @return
	 */
	protected ColPointDto fillStandAloneContent(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelExportCell exportCell, Object value, Row row, CreationHelper createHelper, ExcelExport excelExport, boolean isMap, ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel) {
		//如果是map类型的数据
		if (isMap) {
			return fillStandAloneMap(wb, sheet, colIndex, rowIndex, exportCell, value, createHelper, excelExport, cacheFieldModel);
		} else {
			if (StringUtil.notEmpty(exportCell.titleName())) {
				Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				setCellValueAndStyle(cell, exportCell.titleName() + exportCell.separator(), cacheFieldModel.getTitleStyleName(), null, createHelper);
				ColPointDto pointDto = cellMergedTitle(exportCell, rowIndex, colIndex, sheet);
				colIndex = pointDto.colIndex;
				colIndex++;
			}
			Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
			ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
			value = addValidatorOrComment(wb, sheet, colIndex, rowIndex, exportCell, value);
			if (Objects.nonNull(value) && value instanceof byte[]) {
				createPicture(wb, sheet, cell, (byte[]) value, cacheFieldModel.getContentStyleName());
			} else {
				setCellValueAndStyle(cell, value, cacheFieldModel.getContentStyleName(), exportCell.linkName(), createHelper);
			}

			return cellMerged(exportCell, rowIndex, colIndex, sheet);
		}
	}

	/**
	 * 添加验证和批注
	 *
	 * @param wb
	 * @param sheet
	 * @param colIndex
	 * @param rowIndex
	 * @param exportCell
	 */
	protected Object addValidatorOrComment(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelExportCell exportCell, Object value) {
		// validator
		String[] dropDownOptions = exportCell.dropDownOptions();
		if (Objects.nonNull(value) && value instanceof ComboBoxModel) {
			ComboBoxModel comboBoxModel = ((ComboBoxModel) value);
			createComment(wb, sheet, colIndex, rowIndex, comboBoxModel);
			dropDownOptions = comboBoxModel.getOptions();
			value = comboBoxModel.getValue();
		}
		if (Objects.nonNull(dropDownOptions) && dropDownOptions.length > ExcelConstant.ZERO_SHORT) {
			CellRangeAddressList rangeAddressList = new CellRangeAddressList();
			rangeAddressList.addCellRangeAddress(rowIndex, colIndex, rowIndex, colIndex);
			ExcelHelper.addDropDownValidation(sheet, rangeAddressList, dropDownOptions);
		}
		if (Objects.nonNull(value) && value instanceof NumberScopeModel) {
			NumberScopeModel scopeModel = ((NumberScopeModel) value);
			createComment(wb, sheet, colIndex, rowIndex, scopeModel);
			value = scopeModel.getValue();
			CellRangeAddressList rangeAddressList = new CellRangeAddressList();
			rangeAddressList.addCellRangeAddress(rowIndex, colIndex, rowIndex, colIndex);
			ExcelHelper.addRangeValidation(sheet, rangeAddressList, scopeModel.getStart(), scopeModel.getEnd());
		}
		value = createComment(wb, sheet, colIndex, rowIndex, exportCell, value);
		return value;
	}

	/**
	 * 创建批注
	 *
	 * @param wb
	 * @param sheet
	 * @param colIndex
	 * @param rowIndex
	 * @param exportCell
	 * @param value
	 * @return
	 */
	protected Object createComment(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelExportCell exportCell, Object value) {
		// comment
		String commentText = exportCell.commentText(), commentFontName = exportCell.commentFontName();
		if (Objects.nonNull(value) && value instanceof CommentModel) {
			CommentModel commentModel = ((CommentModel) value);
			commentText = commentModel.getCommentText();
			commentFontName = commentModel.getCommentFontName();
			value = commentModel.getValue();
		}
		if (StringUtil.notEmpty(commentText)) {
			ExcelHelper.createComment(wb, sheet, rowIndex, colIndex, ExcelConstant.NULL_STR, commentText, fontLocal.get().get(commentFontName));
		}
		return value;
	}

	protected void createComment(Workbook wb, Sheet sheet, int colIndex, int rowIndex, CommentModel commentModel) {
		// comment
		String commentText = commentModel.getCommentText(), commentFontName = commentModel.getCommentFontName();
		if (StringUtil.notEmpty(commentText)) {
			ExcelHelper.createComment(wb, sheet, rowIndex, colIndex, ExcelConstant.NULL_STR, commentText, fontLocal.get().get(commentFontName));
		}
	}

	protected ColPointDto fillStandAloneMap(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelExportCell exportCell, Object value, CreationHelper createHelper, ExcelExport excelExport, ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel) {
		final int initColIndex = colIndex;
		final int initRowIndex = rowIndex;
		if (value instanceof Map) {
			Iterator keyIterator = ((Map) value).keySet().iterator();

			while (keyIterator.hasNext()) {
				String title = "";
				if (StringUtil.notEmpty(exportCell.titleName())) {
					title = title + exportCell.titleName();
				}

				String key = keyIterator.next().toString();
				if (Strings.isNullOrEmpty(title)) {
					title = key;
				} else {
					title = title + ExcelConstant.SEPARATOR + key;
				}
				Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
				Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				ExcelHelper.setRowHeight(row, exportCell.rowHeight());
				String styleName = StringUtil.notEmpty(exportCell.titleStyleName()) ? exportCell.titleStyleName() : excelExport.titleStyleName();
				setCellValueAndStyle(cell, title + exportCell.separator(), styleName, null, createHelper);
				ColPointDto titlePointDto = cellMergedTitle(exportCell, rowIndex, colIndex, sheet);
				colIndex = titlePointDto.colIndex;
				rowIndex = titlePointDto.rowIndex;
				colIndex++;

				Cell valueCell = ExcelHelper.getCellOrCreate(row, colIndex);
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				ExcelHelper.setRowHeight(row, exportCell.rowHeight());
				Object val = ((Map) value).get(key);
				ColPointDto pointDto = cellMerged(exportCell, rowIndex, colIndex, sheet);
				rowIndex = pointDto.rowIndex;
				colIndex = pointDto.colIndex;
				if (Objects.nonNull(val) && val instanceof byte[]) {
					createPicture(wb, sheet, cell, (byte[]) val, exportCell.contentStyleName());
				} else {
					String valueStyleName = StringUtil.notEmpty(exportCell.contentStyleName()) ? exportCell.contentStyleName() : excelExport.contentStyleName();
					setCellValueAndStyle(valueCell, val, valueStyleName, exportCell.linkName(), createHelper);
				}

				if (exportCell.fillStyle() == ExcelExportFillStyleEnum.HORIZONTAL) {
					colIndex++;
					rowIndex = initRowIndex;
				} else {
					rowIndex++;
					colIndex = initColIndex;
				}
			}
			Map map = (Map) value;
			if (map.keySet().size() > ExcelConstant.ZERO_SHORT) {
				if (exportCell.fillStyle() == ExcelExportFillStyleEnum.HORIZONTAL) {
					colIndex--;
				} else {
					rowIndex--;
				}
			}

		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}


	/**
	 * 添加导出模型
	 *
	 * @param cla       导出模型class
	 * @param sheetName sheet名称
	 * @param model     单个模型
	 * @param modelList 模型list
	 */
	protected <T extends ExcelBaseModel> void addModel(Class<? extends ExcelBaseModel> cla, String sheetName, T model, List<T> modelList, Integer rowIndex, Integer colIndex, boolean fillTemplate, List<ExcelExportModel> fillList) {
		ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(cla);
		if (Objects.isNull(cacheModel)) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_LOAD);
		}
		ExcelExportModel exportModel = new ExcelExportModel();
		exportModel.setDataModelList(modelList);
		exportModel.setCacheModel(cacheModel);
		exportModel.setSheetName(sheetName);
		exportModel.setDataModel(model);
		exportModel.setExcelModelClass(cla);
		exportModel.setColIndex(colIndex);
		exportModel.setRowIndex(rowIndex);
		exportModel.setFillTemplate(fillTemplate);
		fillList.add(exportModel);
	}

	/**
	 * 设置单元格内容和样式
	 *
	 * @param cell         单元格
	 * @param value        数据
	 * @param styleName    样式名称
	 * @param linkName     连接名称
	 * @param createHelper excelHelper
	 */
	protected void setCellValueAndStyle(Cell cell, Object value, String styleName, String linkName, CreationHelper createHelper) {
		value = setStyle(cell, value, styleName, linkName);
		if (Objects.isNull(value)) {
			return;
		}
		value = ExcelHelper.createHyperlink(cell, value, linkName, createHelper);
		// 设置内容
		ExcelHelper.setCellValue(cell, value);
	}

	/**
	 * 设置单元格内容和样式
	 *
	 * @param cell         单元格
	 * @param value        数据
	 * @param styleName    样式名称
	 * @param linkName     连接名称
	 * @param createHelper excelHelper
	 * @param expression   表达式
	 */
	protected void setCellValueAndStyle(Cell cell, Object value, String styleName, String linkName, CreationHelper createHelper, String expression) {
		value = setStyle(cell, value, styleName, linkName);
		value = Objects.isNull(value) ? ExcelConstant.NULL_STR : value;
		value = ExcelHelper.createHyperlink(cell, value, linkName, createHelper);
		replaceExpression(cell, value, expression);
	}

	/**
	 * 替换标签
	 *
	 * @param cell
	 * @param value
	 * @param expression
	 */
	protected void replaceExpression(Cell cell, Object value, String expression) {
		value = Objects.isNull(value) ? ExcelConstant.NULL_STR : value;
		cell.setCellType(CellType.STRING);
		String expressionValue = cell.getStringCellValue();
		if (expressionValue.matches(expression)) {
			ExcelHelper.setCellValue(cell, value);
		} else {
			String cellValue = expressionValue.replaceAll(expression, value.toString());
			cell.setCellValue(cellValue);
		}
	}


	/**
	 * 设置样式
	 *
	 * @param cell      单元格
	 * @param value     值
	 * @param styleName 样式名
	 * @param linkName  连接名称
	 * @return
	 */
	protected Object setStyle(Cell cell, Object value, String styleName, String linkName) {
		CellStyle cellStyle = null;
		// 设置样式
		if (StringUtil.notEmpty(styleName)) {
			cellStyle = styleLocal.get().get(styleName);
			if (null != cellStyle) {
				cell.setCellStyle(cellStyle);
			}
		}
		if (value == null) {
			return value;
		}
		// 设置默认样式
		if (cellStyle == null) {
			if (value instanceof Date || value instanceof Calendar) {
				cell.setCellStyle(styleLocal.get().get(ExcelBasicStyle.STYLE_DATE_YYYYMMDDHHMMSS));
			} else if (StringUtil.notEmpty(linkName)) {
				cell.setCellStyle(styleLocal.get().get(ExcelBasicStyle.STYLE_HLINK));
			}
		}
		return value;
	}


	/**
	 * 格式化数据
	 *
	 * @param data          数据
	 * @param formatPattern 格式化字符串
	 * @param formatter     格式化器
	 * @return Object
	 */
	protected Object formatValue(Object data, String formatPattern, ExcelWriterDataFormat formatter) {
		if (null == data || StringUtil.isEmpty(formatPattern)) {
			return data;
		}
		if (null != formatter) {
			return formatter.format(data, formatPattern);
		}
		return dataFormat.format(data, formatPattern);
	}


	/**
	 * 创建图片
	 *
	 * @param wb    WorkBook
	 * @param sheet sheet
	 * @param cell  单元格
	 * @param bytes 图片流数组
	 * @return Picture
	 */
	protected Picture createPicture(Workbook wb, Sheet sheet, Cell cell, byte[] bytes, String styleName) {
		// 设置样式
		if (StringUtil.notEmpty(styleName)) {
			CellStyle cellStyle = styleLocal.get().get(styleName);
			if (null != cellStyle) {
				cell.setCellStyle(cellStyle);
			}
		}
		int i = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		// 创建锚点
		CreationHelper creationHelper = wb.getCreationHelper();
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setRow1(cell.getRowIndex());
		clientAnchor.setCol1(cell.getColumnIndex());
		clientAnchor.setRow2(cell.getRowIndex() + ExcelConstant.ONE_INT);
		clientAnchor.setCol2(cell.getColumnIndex() + ExcelConstant.ONE_INT);

		Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
		return drawingPatriarch.createPicture(clientAnchor, i);
	}

	/**
	 * @Description: excel 排除处理器
	 * @Author: Vachel Wang
	 * @Date: 2019/12/7 上午9:36
	 * @Email:
	 */
	public class ExcludeHandler {
		private Class<? extends ExcelBaseModel> modelCla;

		public ExcludeHandler(Class<? extends ExcelBaseModel> modelCla) {
			this.modelCla = modelCla;
		}

		/**
		 * 排除
		 *
		 * @param excludeFields
		 */
		public ExcludeHandler excludes(String[] excludeFields) {
			if (Objects.nonNull(excludeFields) && excludeFields.length > ExcelConstant.ZERO_SHORT) {
				Map<String, String> fieldMap = excludeFieldMap.get(modelCla);
				boolean exists = true;
				if (Objects.isNull(fieldMap)) {
					fieldMap = new HashMap<>();
					exists = false;
				}
				for (String field : excludeFields) {
					fieldMap.put(field, ExcelConstant.NULL_STR);
				}
				if (!exists) {
					excludeFieldMap.put(modelCla, fieldMap);
				}
			}
			return this;
		}

		/**
		 * 添加验证或批注
		 *
		 * @param field
		 * @param commentModel
		 * @return
		 */
		public ExcludeHandler addValidationOrComment(String field, CommentModel commentModel) {
			if (Objects.isNull(commentModel) || StringUtil.isEmpty(field)) {
				return this;
			}
			Map<String, CommentModel> fieldMap = commentMap.get(modelCla);
			boolean exists = true;
			if (Objects.isNull(fieldMap)) {
				fieldMap = new HashMap<>();
				exists = false;
			}
			fieldMap.put(field, commentModel);
			if (!exists) {
				commentMap.put(modelCla, fieldMap);
			}
			return this;
		}
	}

	/**
	 * 列坐标
	 */
	@Data
	@Builder
	static class ColPointDto implements Serializable {
		private int colIndex;
		private int rowIndex;
	}

	/**
	 * 目标列表结构dto
	 */
	@Data
	@Builder
	static class ExcelListTemplateDto {
		private Method getMethod;
		private int rowIndex;
		private int colIndex;
		private String sheetName;
		private ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel;
		private ExcelExportCell exportCell;
	}

	/**
	 * 添加空数据默认内容
	 *
	 * @param wb             workbook
	 * @param creationHelper creationHelper
	 */
	protected void addNoResultData(Workbook wb, CreationHelper creationHelper) {
		Sheet sheet = wb.getNumberOfSheets() == ExcelConstant.ZERO_SHORT ? wb.createSheet(ExcelConstant.DEFAULT_SHEET_NAME) : wb.getSheetAt(ExcelConstant.ZERO_SHORT);
		if (sheet.getLastRowNum() != ExcelConstant.ZERO_SHORT) {
			return;
		}
		if (Objects.isNull(listCla)) {
			ExcelCustomColumnModel columnModel = new ExcelCustomColumnModel();
			columnModel.setRowIndex(ExcelConstant.ZERO_SHORT);
			columnModel.setColIndex(ExcelConstant.ZERO_SHORT);
			columnModel.setSheetName(ExcelConstant.DEFAULT_SHEET_NAME);
			if(noneDataTips) {
				columnModel.setColWidth(ExcelConstant.SHORT_500);
				columnModel.setRowHeight(ExcelConstant.SHORT_50);
				columnModel.setStyleName(ExcelBasicStyle.STYLE_ZEBRA_TITLE_ROW);
				columnModel.setValue(ExcelErrorMsgConstant.ERROR_EXPORT_NOT_FOUND_DATA);
			}
			fillColumn(wb, creationHelper, columnModel, columnModel.getRowIndex(), columnModel.getColIndex());
		} else {
			ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(listCla);
			Row row = ExcelHelper.getRowOrCreate(sheet, ExcelConstant.ZERO_SHORT);
			int colIndex = ExcelConstant.ZERO_SHORT;
			for (ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel : cacheModel.getFieldModelList()) {
				Object titleName = cacheFieldModel.getExportCell().titleName();
				Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				ExcelHelper.setRowHeight(row, cacheFieldModel.getTitleRowHeight());

				this.setCellValueAndStyle(cell, titleName, cacheFieldModel.getTitleStyleName(), null, creationHelper);
				colIndex++;
			}
		}
	}
}
