package com.github.excel.read.handler.impl;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.model.ExcelReadErrorMsgInfo;
import com.github.excel.model.ReadPictureModel;
import com.github.excel.read.AbstractExcelReader;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelReadPictureTypeEnum;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelImportTemplateCacheModel;
import com.github.excel.model.ExcelReadErrorMsgModel;
import com.github.excel.util.StringUtil;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:11 下午
 * @Description: 用户模式下解析bean处理器
 */
@Slf4j
public class ExcelUserParseHandler extends AbstractExcelParseHandler {

	public ExcelUserParseHandler(InputStream readExcelStream, String excelName) {
		super(readExcelStream, excelName);
	}

	private void validateTemplate(String template, Workbook workbook, boolean endReturnException) {
		Map<String, List<ExcelImportTemplateCacheModel>> templateCacheMap = ExcelBootLoader.getExcelImportTemplateCacheMapValue(template);
		if (Objects.nonNull(templateCacheMap)) {
			// 校验名称
			String suffix = template.substring(template.indexOf(ExcelConstant.DOT_CHAR));
			if (!suffix.equals(excelName.substring(excelName.indexOf(ExcelConstant.DOT_CHAR)))) {
				addOrThrowError(endReturnException,null,null,String.format(ExcelErrorMsgConstant.ERROR_IMPORT_FILE_SUFFIX,suffix));
			}
			for (Map.Entry<String, List<ExcelImportTemplateCacheModel>> entry : templateCacheMap.entrySet()) {
				Sheet sheet = workbook.getSheet(entry.getKey());
				if (Objects.isNull(sheet)) {
					addOrThrowError(endReturnException, sheet, null, String.format(ExcelErrorMsgConstant.ERROR_NOT_FOUND_SHEET, entry.getKey()));
					continue;
				}
				List<ExcelImportTemplateCacheModel> cacheModelList = entry.getValue();
				for (ExcelImportTemplateCacheModel cacheModel : cacheModelList) {
					Row row = sheet.getRow(cacheModel.getRowIndex());
					if (Objects.isNull(row)) {
						addOrThrowError(endReturnException, sheet, null, String.format(ExcelErrorMsgConstant.ERROR_NOT_FOUND_SHEET_ROW, entry.getKey(), String.valueOf(cacheModel.getRowIndex() + ExcelConstant.ONE_INT)));
						continue;
					}
					Cell cell = row.getCell(cacheModel.getColIndex());
					if (Objects.isNull(cell)) {
						addOrThrowErrorByCell(endReturnException, sheet, row, String.format(ExcelErrorMsgConstant.ERROR_NOT_FOUND_SHEET_ROW_COL, entry.getKey(), String.valueOf(cacheModel.getRowIndex() + ExcelConstant.ONE_INT), String.valueOf(cacheModel.getColIndex() + ExcelConstant.ONE_INT)));
						continue;
					}
					cell.setCellType(CellType.STRING);
					String value = cell.getStringCellValue();
					if (StringUtil.isEmpty(value) || !value.equals(cacheModel.getText())) {
						String errorTips = String.format(ExcelErrorMsgConstant.ERROR_NOT_MATCH_COL_CONTENT, entry.getKey(), cell.getAddress().formatAsString(), cacheModel.getText());
						addOrThrowError(endReturnException, sheet, cell, null, errorTips);
					}
				}
			}
		}
	}

	@Override
	public ExcelReadErrorMsgModel process(List<AbstractExcelReader.ExcelReadModelDto> readModelDtoList, boolean endReturnException, boolean readPicture, String template) {

		Map<Integer, List<AbstractExcelReader.ExcelReadModelDto>> sheetModelMap = readModelDtoList.stream().collect(Collectors.groupingBy(AbstractExcelReader.ExcelReadModelDto::getSheetIndex));

		try (Workbook workbook = excelName.endsWith(ExcelConstant.XLSX_STR) ? new XSSFWorkbook(readExcelStream) : new HSSFWorkbook(readExcelStream)) {
			// 校验模板
			if (StringUtil.notEmpty(template)) {
				validateTemplate(template, workbook, endReturnException);
				if (this.errorMsgModel.getExistsError()) {
					return this.errorMsgModel;
				}
			}
			for (Map.Entry<Integer, List<AbstractExcelReader.ExcelReadModelDto>> readModelEntry : sheetModelMap.entrySet()) {
				Sheet sheet = workbook.getSheetAt(readModelEntry.getKey());
				Map<String, List<ReadPictureModel>> sheetPictureMap = Maps.newHashMap();
				if (readPicture) {
					if (workbook instanceof XSSFWorkbook) {
						sheetPictureMap = getXSSFSheetPicture((XSSFSheet) sheet);
					} else {
						sheetPictureMap = getHSSFSheetPicture((HSSFSheet) sheet);
					}
				}

				List<AbstractExcelReader.ExcelReadModelDto> readModelList = readModelEntry.getValue();
				if (Objects.isNull(sheet)) {
					continue;
				}

				for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
					Row row = sheet.getRow(rowIndex);
					if (Objects.isNull(row)) {
						continue;
					}
					for (int cellIndex = row.getFirstCellNum(); cellIndex < row.getLastCellNum(); cellIndex++) {
						Cell cell = row.getCell(cellIndex);
						if (Objects.isNull(cell)) {
							continue;
						}
						cell.setCellType(CellType.STRING);
						String titleName = cell.getStringCellValue();
						if (StringUtil.isEmpty(titleName)) {
							continue;
						}
						Cell dataCell = row.getCell(cellIndex + ExcelConstant.ONE_INT);

						// 解析成bean
						List<AbstractExcelReader.ExcelReadModelDto> fillBeanDtoList = readModelList.stream().filter(e -> {
							return (e.getCacheImportModel().getFieldModelMap().get(titleName) != null && e.getModelList() == null);
						}).collect(Collectors.toList());

						if (fillBeanDtoList.size() > ExcelConstant.ZERO_SHORT) {
							cellIndex++;
							for (AbstractExcelReader.ExcelReadModelDto modelDto : fillBeanDtoList) {
								ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = modelDto.getCacheImportModel().getFieldModelMap().get(titleName);
								Class<?> parameterType = fieldModel.getSetMethod().getParameterTypes()[ExcelConstant.ZERO_SHORT];
								Object setParams = getCellValue(dataCell, parameterType);
								// get picture value
								setParams = getPictureValue(readPicture, sheetPictureMap, rowIndex, cellIndex, setParams, parameterType);

								// 校验为空数据
								if (fieldModel.getImportProperty().checkNull() && Objects.isNull(setParams)) {
									CellReference cellReference = new CellReference(rowIndex, cellIndex);
									addNullErrorMsgOrThrow(endReturnException, sheet, cellReference, fieldModel, readModelEntry.getKey());
									continue;
								}

								if (Objects.isNull(setParams)) {
									continue;
								}
								// validation number format
								if (Number.class.isAssignableFrom(parameterType) && setParams instanceof String) {
									if (!setParams.toString().matches(ExcelConstant.NUMBER_PATTERN)) {
										addFormatErrorMsgOrThrow(endReturnException, sheet, dataCell, fieldModel, readModelEntry.getKey());
										continue;
									}
								}
								// 格式化数据
								try {
									setParams = getDataFormatThenCache(fieldModel.getImportProperty().formatter()).format(setParams, fieldModel.getImportProperty().formatPattern(), parameterType);
								} catch (ParseException e) {
									log.error(Throwables.getStackTraceAsString(e));
									addFormatErrorMsgOrThrow(endReturnException, sheet, dataCell, fieldModel, readModelEntry.getKey());
									continue;
								}
								// 填充数据
								try {
									fieldModel.getSetMethod().invoke(modelDto.getModel(), setParams);
								} catch (IllegalAccessException e) {
									addOrThrowError(endReturnException, sheet, dataCell, readModelEntry.getKey(), String.format(ExcelErrorMsgConstant.ERROR_DATA_INVOKE_READ_MSG, sheet.getSheetName(), dataCell.getAddress().formatAsString(), fieldModel.getImportProperty().titleName()));
								} catch (IllegalArgumentException e) {
									addOrThrowError(endReturnException, sheet, dataCell, readModelEntry.getKey(), String.format(ExcelErrorMsgConstant.ERROR_DATA_INVOKE_READ_MSG, sheet.getSheetName(), dataCell.getAddress().formatAsString(), fieldModel.getImportProperty().titleName()));
								} catch (InvocationTargetException e) {
									addOrThrowError(endReturnException, sheet, dataCell, readModelEntry.getKey(), String.format(ExcelErrorMsgConstant.ERROR_DATA_INVOKE_READ_MSG, sheet.getSheetName(), dataCell.getAddress().formatAsString(), fieldModel.getImportProperty().titleName()));
								}
							}
						} else {
							// 解析成bean list
							Cell titleCell2 = row.getCell(cellIndex + ExcelConstant.ONE_INT);
							if (Objects.isNull(titleCell2)) {
								continue;
							}
							titleCell2.setCellType(CellType.STRING);
							String titleName2 = titleCell2.getStringCellValue();
							if (StringUtil.isEmpty(titleCell2)) {
								continue;
							}
							// 通过title获取两列，如果都可以同时获取到则映射为title
							fillBeanDtoList = readModelList.stream().filter(e -> {
								return (e.getCacheImportModel().getFieldModelMap().get(titleName) != null && e.getCacheImportModel().getFieldModelMap().get(titleName2) != null && e.getModelList() != null);

							}).collect(Collectors.toList());
							if (!fillBeanDtoList.isEmpty()) {
								rowIndex = parseList(fillBeanDtoList, sheet, rowIndex, cellIndex, endReturnException, readModelEntry.getKey(), readPicture, sheetPictureMap);
								break;
							}
						}

					}
				}
				// 执行校验
				for (AbstractExcelReader.ExcelReadModelDto modelDto : readModelList) {
					if (modelDto.getModelList() == null) {
						validateBean(sheet.getSheetName(), null, modelDto.getModel(), endReturnException);
					}
				}
			}

			return errorMsgModel;
		} catch (IOException e) {
			log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException(e.getMessage());
		} catch (IllegalAccessException e) {
			log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException(e.getMessage());
		} catch (InvocationTargetException e) {
			log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException(e.getMessage());
		} catch (InstantiationException e) {
			log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException(e.getMessage());
		}
	}

	/**
	 * 校验数据
	 *
	 * @param baseModel
	 * @param endReturnException
	 */
	private void validateBean(String sheetName, Integer rowIndex, ExcelBaseModel baseModel, boolean endReturnException) {
		try {
			baseModel.validator();
		} catch (ExcelReadException ex) {
			if (endReturnException) {
				ExcelReadErrorMsgInfo errorMsgInfo = new ExcelReadErrorMsgInfo(ex.getMessage());
				errorMsgModel.getErrorMsgInfoList().add(errorMsgInfo);
				errorMsgModel.setExistsError(true);
			} else {
				throw ex;
			}
		}
	}

	/**
	 * 映射标题
	 *
	 * @param sheet
	 * @param startRow
	 * @param startCol
	 * @param fillBeanDtoList
	 * @return
	 */
	private mappingTitleDto mappingTitle(Sheet sheet, int startRow, int startCol, List<AbstractExcelReader.ExcelReadModelDto> fillBeanDtoList) {
		Row titleRow = sheet.getRow(startRow);
		Map<Integer, List<ExcelReadMappingDto>> readMappingMap = Maps.newHashMap();
		int endColIndex = titleRow.getLastCellNum();
		// 先确定实体titleName与Excel标题列的映射关系，通过 Map<Integer, List<ExcelReadMappingDto>> 来确定，key 为列号，Value 表示每一个字段，每一个字段对应有多个需要解析的对象
		// 注意：这里解析的是一行title的名称，不能为合并单元格，否则不会解析下一个相邻的标题单元格
		for (int colIndex = startCol; colIndex < titleRow.getLastCellNum(); colIndex++) {
			Cell cell = titleRow.getCell(colIndex);
			// 确定标题最后一列坐标
			if (Objects.isNull(cell)) {
				endColIndex = colIndex;
				break;
			}
			cell.setCellType(CellType.STRING);
			String title = cell.getStringCellValue();
			if (StringUtil.isEmpty(title)) {
				endColIndex = colIndex;
				break;
			}
			List<ExcelReadMappingDto> readMapingDtoList = Lists.newArrayList();
			for (AbstractExcelReader.ExcelReadModelDto readModelDto : fillBeanDtoList) {
				ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = readModelDto.getCacheImportModel().getFieldModelMap().get(title);
				// 没有这个字段就不设值
				if (Objects.isNull(fieldModel)) {
					continue;
				}
				ExcelReadMappingDto mapingDto = new ExcelReadMappingDto(fieldModel, readModelDto);
				readMapingDtoList.add(mapingDto);
			}
			readMappingMap.put(colIndex, readMapingDtoList);
		}
		startRow++;
		return new mappingTitleDto(startRow, endColIndex, readMappingMap);
	}


	/**
	 * 解析Excel成list
	 *
	 * @param fillBeanDtoList
	 * @param sheet
	 * @param startRow
	 * @param startCol
	 * @param endReturnException
	 * @param sheetIndex
	 * @param readPicture
	 * @param sheetPictureMap
	 * @return
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws InvocationTargetException
	 */
	private int parseList(List<AbstractExcelReader.ExcelReadModelDto> fillBeanDtoList, Sheet sheet, int startRow, int startCol, boolean endReturnException, Integer sheetIndex, boolean readPicture, Map<String, List<ReadPictureModel>> sheetPictureMap) throws IllegalAccessException, InstantiationException, InvocationTargetException {
		final int initStartCol = startCol;

		mappingTitleDto mappingTitleDto = mappingTitle(sheet, startRow, startCol, fillBeanDtoList);
		return parseListContent(fillBeanDtoList, sheet, endReturnException, sheetIndex, initStartCol, mappingTitleDto, readPicture, sheetPictureMap);
	}

	/**
	 * 解析内容
	 *
	 * @param fillBeanDtoList    填充bean
	 * @param sheet              sheet
	 * @param endReturnException 异常磨损
	 * @param sheetIndex         sheet下标
	 * @param initStartCol       初始化列号
	 * @param mappingTitleDto    标题映射
	 * @param readPicture
	 * @param sheetPictureMap
	 * @return
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	private int parseListContent(List<AbstractExcelReader.ExcelReadModelDto> fillBeanDtoList, Sheet sheet, boolean endReturnException, Integer sheetIndex, final int initStartCol, mappingTitleDto mappingTitleDto, boolean readPicture, Map<String, List<ReadPictureModel>> sheetPictureMap) throws InstantiationException, IllegalAccessException, InvocationTargetException {
		int startRow;
		startRow = mappingTitleDto.getStartRow();
		int endColIndex = mappingTitleDto.getEndColIndex();
		Map<Integer, List<ExcelReadMappingDto>> readMappingMap = mappingTitleDto.getReadMappingMap();

		int rowIndex = startRow;
		// 开始解析内容。解析出来的映射来设值，获取映射的方式是通过下标（mappingIndex）来获取
		for (; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			// 判断列表解析完成是通过行为空来判断
			Row row = sheet.getRow(rowIndex);
			if (Objects.isNull(row)) {
				break;
			}

			// 为多个需要解析的List，先创建Element。通过class来映射需要填充的对象
			Map<Class<? extends ExcelBaseModel>, ListElement> elementMap = Maps.newHashMap();
			for (AbstractExcelReader.ExcelReadModelDto readModelDto : fillBeanDtoList) {
				Object element = readModelDto.getModelCla().newInstance();
				ListElement listElement = new ListElement(element,readModelDto.getModelList(),false);
				elementMap.put(readModelDto.getModelCla(), listElement);
//				readModelDto.getModelList().add(element);
			}
			// 错误统计
			List<ReadExceptionInfo> exceptionInfoList = new ArrayList<>();
			// 是否填充标识
			boolean hasFill = false;
			for (int colIndex = initStartCol; colIndex < endColIndex; colIndex++) {
				List<ExcelReadMappingDto> mappingDtoList = readMappingMap.get(colIndex);
				if (mappingDtoList.isEmpty()) {
					// 没有需要解析的则跳过
					continue;
				}
				// 解析每一个字段
				Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

				// 决定哪些东西需要填充
				for (ExcelReadMappingDto mappingDto : mappingDtoList) {
					ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = mappingDto.getFieldModel();
					Class<?> parameterType = fieldModel.getSetMethod().getParameterTypes()[ExcelConstant.ZERO_SHORT];
					Object setParams = Objects.isNull(cell) ? null : getCellValue(cell, parameterType);

					// get picture value
					setParams = getPictureValue(readPicture, sheetPictureMap, rowIndex, colIndex, setParams, parameterType);

					if (fieldModel.getImportProperty().checkNull() && Objects.isNull(setParams)) {
						CellReference cellReference = new CellReference(rowIndex, colIndex);
						ReadExceptionInfo readExceptionInfo = createReadExceptionInfo(endReturnException, sheet, cellReference, fieldModel, sheetIndex);
						exceptionInfoList.add(readExceptionInfo);
						continue;
					}
					if (Objects.isNull(setParams)) {
						continue;
					}
					// validation number format
					if (Number.class.isAssignableFrom(parameterType)) {
						if (!setParams.toString().matches(ExcelConstant.NUMBER_PATTERN)) {
							addFormatErrorMsgOrThrow(endReturnException, sheet, cell, fieldModel, sheetIndex);
							continue;
						}
					}
					try {
						setParams = getDataFormatThenCache(fieldModel.getImportProperty().formatter()).format(setParams, fieldModel.getImportProperty().formatPattern(), parameterType);
					} catch (ParseException e) {
						log.error(Throwables.getStackTraceAsString(e));
						addFormatErrorMsgOrThrow(endReturnException, sheet, cell, fieldModel, sheetIndex);
						continue;
					}

					try {
						ListElement listElement = elementMap.get(mappingDto.getReadModelDto().getModelCla());
						fieldModel.getSetMethod().invoke(listElement.getObj(), setParams);
						listElement.setHasFilled(true);
						hasFill = true;
					} catch (IllegalAccessException e) {
						addOrThrowError(endReturnException, sheet, cell, sheetIndex, String.format(ExcelErrorMsgConstant.ERROR_DATA_INVOKE_READ_MSG, sheet.getSheetName(), cell.getAddress().formatAsString(), fieldModel.getImportProperty().titleName()));
					} catch (IllegalArgumentException e) {
						addOrThrowError(endReturnException, sheet, cell, sheetIndex, String.format(ExcelErrorMsgConstant.ERROR_DATA_INVOKE_READ_MSG, sheet.getSheetName(), cell.getAddress().formatAsString(), fieldModel.getImportProperty().titleName()));
					} catch (InvocationTargetException e) {
						addOrThrowError(endReturnException, sheet, cell, sheetIndex, String.format(ExcelErrorMsgConstant.ERROR_DATA_INVOKE_READ_MSG, sheet.getSheetName(), cell.getAddress().formatAsString(), fieldModel.getImportProperty().titleName()));
					}
				}
			}
			elementMap.forEach((k,v)->{
				if (v.hasFilled) {
					ExcelBaseModel baseModel = (ExcelBaseModel) v.getObj();
					v.getModelList().add(baseModel);
				}
			});
			// 没有填充
			if (!hasFill) {
				break;
			} else {
				for (ReadExceptionInfo exceptionInfo : exceptionInfoList) {
					addOrThrowError(exceptionInfo.endReturnException, exceptionInfo.sheet, exceptionInfo.cellReference, exceptionInfo.sheetIndex, exceptionInfo.errorTips);
				}
			}
			// validator and batch processing
			for (AbstractExcelReader.ExcelReadModelDto readModelDto : fillBeanDtoList) {
				ExcelBaseModel baseModel = (ExcelBaseModel) readModelDto.getModelList().get(readModelDto.getModelList().size() - ExcelConstant.ONE_INT);
				validateBean(sheet.getSheetName(), row.getRowNum(), baseModel, endReturnException);
				// batch processing
				if (readModelDto.getBatchExecute()) {
					try {
						readModelDto.getBatchProcess().doProcess(readModelDto.getModelList());
					} catch (ExcelReadException e) {
						addOrThrowError(endReturnException, sheet, sheetIndex, e.getMessage());
					}

				}
			}
		}
		// end processing
		for (AbstractExcelReader.ExcelReadModelDto readModelDto : fillBeanDtoList) {
			// batch processing
			if (readModelDto.getBatchExecute() && !readModelDto.getModelList().isEmpty()) {
				try {
					readModelDto.getBatchProcess().process(readModelDto.getModelList());
					readModelDto.getModelList().clear();
				} catch (ExcelReadException e) {
					addOrThrowError(endReturnException, sheet, sheetIndex, e.getMessage());
				}

			}
		}
		return rowIndex;
	}

	/**
	 * @Description: 获取图片内容
	 * @Param:
	 * @return:
	 * @Author: Vachel Wang
	 * @Date: 2019/10/29 下午3:03
	 */
	private Object getPictureValue(boolean readPicture, Map<String, List<ReadPictureModel>> sheetPictureMap, int rowIndex, int colIndex, Object setParams, Class<?> parameterType) {
		if (ReadPictureModel.class.isAssignableFrom(parameterType) && readPicture) {
			List<ReadPictureModel> pictureModelList = sheetPictureMap.get(String.valueOf(rowIndex) + ExcelConstant.COMMA_CHAR + colIndex);
			if (!pictureModelList.isEmpty()) {
				setParams = pictureModelList.get(ExcelConstant.ZERO_SHORT);
			}
		} else if (List.class.isAssignableFrom(parameterType) && readPicture) {
			setParams = sheetPictureMap.get(String.valueOf(rowIndex) + ExcelConstant.COMMA_CHAR + colIndex);
		}
		return setParams;
	}

	/**
	 * @Description: 实体映射dto
	 * @Author: Vachel Wang
	 * @Date: 2019/10/23 下午5:36
	 */
	@Data
	private static class ExcelReadMappingDto<T extends ExcelBaseModel> {
		public ExcelReadMappingDto(ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel, AbstractExcelReader.ExcelReadModelDto readModelDto) {
			this.fieldModel = fieldModel;
			this.readModelDto = readModelDto;
		}

		private ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel;
		private AbstractExcelReader.ExcelReadModelDto readModelDto;
	}


	/**
	 * @Description: 获取单元格值
	 * @Param:
	 * @return:
	 * @Author: Vachel Wang
	 * @Date: 2019/10/23 下午4:54
	 */
	public static Object getCellValue(Cell cell, Class<?> paramterType) {
		Object valueObj = null;
		if (Objects.isNull(cell)) return valueObj;
		if (String.class.isAssignableFrom(paramterType) && cell.getCellTypeEnum() != CellType.STRING) {
			cell.setCellType(CellType.STRING);
		}
		switch (cell.getCellTypeEnum()) {
			case STRING:
				valueObj = cell.getRichStringCellValue().getString().trim();
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					valueObj = cell.getDateCellValue();
				} else {
					valueObj = cell.getNumericCellValue();
				}
				break;
			case BOOLEAN:
				valueObj = String.valueOf(cell.getBooleanCellValue());
				break;
			case FORMULA:
				FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
				evaluator.evaluateInCell(cell);
				CellValue cellValue = evaluator.evaluate(cell);
				if (StringUtil.notEmpty(cellValue.getStringValue())) {
					valueObj = cellValue.getStringValue();
				} else {
					valueObj = cellValue.getNumberValue();
				}
				break;
			default:
				valueObj = null;
		}
		return valueObj;
	}

	/**
	 * 添加或抛出格式化错误
	 *
	 * @param endReturnException
	 * @param sheet
	 * @param cell
	 * @param fieldModel
	 * @param sheetIndex
	 */
	private void addFormatErrorMsgOrThrow(boolean endReturnException, Sheet sheet, Cell cell, ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel, Integer sheetIndex) {
		String errorTips = String.format(ExcelErrorMsgConstant.ERROR_DATA_FORMAT_READ_MSG, sheet.getSheetName(), cell.getAddress().formatAsString(), fieldModel.getImportProperty().titleName());
		addOrThrowError(endReturnException, sheet, cell, sheetIndex, errorTips);
	}

	/**
	 * 添加或抛出数据为空错误
	 *
	 * @param endReturnException
	 * @param sheet
	 * @param cellReference
	 * @param fieldModel
	 * @param sheetIndex
	 */
	private void addNullErrorMsgOrThrow(boolean endReturnException, Sheet sheet, CellReference cellReference, ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel, Integer sheetIndex) {
		String errorTips = String.format(ExcelErrorMsgConstant.ERROR_DATA_NULL_READ_MSG, sheet.getSheetName(), cellReference.formatAsString(), fieldModel.getImportProperty().titleName());
		addOrThrowError(endReturnException, sheet, cellReference, sheetIndex, errorTips);
	}

	/**
	 * 创建错误信息
	 *
	 * @param endReturnException
	 * @param sheet
	 * @param cellReference
	 * @param fieldModel
	 * @param sheetIndex
	 * @return
	 */
	private ReadExceptionInfo createReadExceptionInfo(boolean endReturnException, Sheet sheet, CellReference cellReference, ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel, Integer sheetIndex) {
		String errorTips = String.format(ExcelErrorMsgConstant.ERROR_DATA_NULL_READ_MSG, sheet.getSheetName(), cellReference.formatAsString(), fieldModel.getImportProperty().titleName());
		ReadExceptionInfo readExceptionInfo = new ReadExceptionInfo(endReturnException, sheet, cellReference, sheetIndex, errorTips);
		return readExceptionInfo;
	}

	/**
	 * 添加或抛出错误
	 *
	 * @param endReturnException
	 * @param sheet
	 * @param cell
	 * @param sheetIndex
	 * @param errorTips
	 */
	private void addOrThrowError(boolean endReturnException, Sheet sheet, Cell cell, Integer sheetIndex, String errorTips) {
		if (endReturnException) {
			ExcelReadErrorMsgInfo errorMsgInfo = new ExcelReadErrorMsgInfo(sheetIndex, sheet.getSheetName(), cell.getRowIndex(), cell.getColumnIndex(), cell.getAddress().formatAsString(), errorTips);
			errorMsgModel.getErrorMsgInfoList().add(errorMsgInfo);
			errorMsgModel.setExistsError(true);
		} else {
			throw new ExcelReadException(errorTips);
		}
	}

	/**
	 * 添加或抛出异常
	 *
	 * @param endReturnException
	 * @param sheet
	 * @param sheetIndex
	 * @param errorTips
	 */
	private void addOrThrowError(boolean endReturnException, Sheet sheet, Integer sheetIndex, String errorTips) {
		if (endReturnException) {
			String sheetName = Objects.isNull(sheet) ? null : sheet.getSheetName();
			ExcelReadErrorMsgInfo errorMsgInfo = new ExcelReadErrorMsgInfo(sheetIndex, sheetName, null, null, null, errorTips);
			errorMsgModel.getErrorMsgInfoList().add(errorMsgInfo);
			errorMsgModel.setExistsError(true);
		} else {
			throw new ExcelReadException(errorTips);
		}
	}

	/**
	 * 添加或抛出异常
	 *
	 * @param endReturnException
	 * @param sheet
	 * @param row
	 * @param errorTips
	 */
	private void addOrThrowErrorByCell(boolean endReturnException, Sheet sheet, Row row, String errorTips) {
		if (endReturnException) {
			ExcelReadErrorMsgInfo errorMsgInfo = new ExcelReadErrorMsgInfo(null, sheet.getSheetName(), row.getRowNum(), null, null, errorTips);
			errorMsgModel.getErrorMsgInfoList().add(errorMsgInfo);
			errorMsgModel.setExistsError(true);
		} else {
			throw new ExcelReadException(errorTips);
		}
	}

	/**
	 * 添加或抛出错误
	 *
	 * @param endReturnException
	 * @param sheet
	 * @param cellReference
	 * @param sheetIndex
	 * @param errorTips
	 */
	private void addOrThrowError(boolean endReturnException, Sheet sheet, CellReference cellReference, Integer sheetIndex, String errorTips) {
		if (endReturnException) {
			ExcelReadErrorMsgInfo errorMsgInfo = new ExcelReadErrorMsgInfo(sheetIndex, sheet.getSheetName(), cellReference.getRow(), (int) cellReference.getCol(), cellReference.formatAsString(), errorTips);
			errorMsgModel.getErrorMsgInfoList().add(errorMsgInfo);
			errorMsgModel.setExistsError(true);
		} else {
			throw new ExcelReadException(errorTips);
		}
	}

	/**
	 * 获取xlsx 图片
	 *
	 * @param sheet
	 * @return
	 */
	public Map<String, List<ReadPictureModel>> getXSSFSheetPicture(XSSFSheet sheet) {
		//returns the existing SpreadsheetDrawingML from the sheet, or creates a new one
		XSSFDrawing drawing = sheet.createDrawingPatriarch();
		//loop through all of the shapes in the drawing area
		List<ReadPictureModel> pictureModelList = Lists.newArrayList();
		for (XSSFShape shape : drawing.getShapes()) {
			if (shape instanceof Picture) {
				//convert the shape into a picture
				XSSFPicture picture = (XSSFPicture) shape;
				XSSFClientAnchor clientAnchor = picture.getClientAnchor();
				String suffix = ExcelReadPictureTypeEnum.getTypeSuffix(picture.getPictureData().getPictureType());
				ReadPictureModel pictureModel = new ReadPictureModel();
				pictureModel.setBytes(picture.getPictureData().getData());
				pictureModel.setRowIndex(clientAnchor.getRow1());
				pictureModel.setColIndex((int) clientAnchor.getCol1());
				pictureModel.setSuffix(suffix);
				pictureModel.setSheetName(sheet.getSheetName());
				pictureModel.setPoint(String.valueOf(clientAnchor.getRow1()) + ExcelConstant.COMMA_CHAR + clientAnchor.getCol1());
				pictureModelList.add(pictureModel);
			}
		}
		return pictureModelList.stream().collect(Collectors.groupingBy(ReadPictureModel::getPoint));
	}

	/**
	 * 获取xls 图片
	 *
	 * @param sheet
	 * @return
	 */
	public Map<String, List<ReadPictureModel>> getHSSFSheetPicture(HSSFSheet sheet) {
		List<ReadPictureModel> pictureModelList = Lists.newArrayList();
		List<HSSFShape> shapes = sheet.getDrawingPatriarch().getChildren();
		for (HSSFShape shape : shapes) {
			if (shape instanceof HSSFPicture) {
				HSSFPicture pic = (HSSFPicture) shape;
				HSSFClientAnchor clientAnchor = pic.getClientAnchor();
				HSSFPictureData picData = sheet.getWorkbook().getAllPictures().get(pic.getPictureIndex() - ExcelConstant.ONE_INT);
				String suffix = ExcelReadPictureTypeEnum.getTypeSuffix(picData.getPictureType());
				ReadPictureModel pictureModel = new ReadPictureModel();
				pictureModel.setBytes(picData.getData());
				pictureModel.setRowIndex(clientAnchor.getRow1());
				pictureModel.setColIndex((int) clientAnchor.getCol1());
				pictureModel.setSuffix(suffix);
				pictureModel.setSheetName(sheet.getSheetName());
				pictureModel.setPoint(String.valueOf(clientAnchor.getRow1()) + ExcelConstant.COMMA_CHAR + clientAnchor.getCol1());
				pictureModelList.add(pictureModel);
			}
		}
		return pictureModelList.stream().collect(Collectors.groupingBy(ReadPictureModel::getPoint));
	}

	/**
	 * 标题映射结果
	 */
	@Data
	private static class mappingTitleDto {
		private int startRow;
		private int endColIndex;
		Map<Integer, List<ExcelReadMappingDto>> readMappingMap;

		public mappingTitleDto(int startRow, int endColIndex, Map<Integer, List<ExcelReadMappingDto>> readMappingMap) {
			this.startRow = startRow;
			this.endColIndex = endColIndex;
			this.readMappingMap = readMappingMap;
		}
	}

	/**
	 * 错误信息
	 */
	@Data
	@AllArgsConstructor
	private class ReadExceptionInfo {
		private Boolean endReturnException;
		private Sheet sheet;
		private CellReference cellReference;
		private Integer sheetIndex;
		private String errorTips;
	}
	/**
	 * 列表填充元素
	 */
	@Data
	@AllArgsConstructor
	private class ListElement<T extends ExcelBaseModel> {
		private Object obj ;
		private List<T> modelList;
		private Boolean hasFilled;
	}
}
