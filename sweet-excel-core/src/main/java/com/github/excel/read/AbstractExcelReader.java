package com.github.excel.read;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.model.ExcelCacheImportModel;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReadErrorMsgModel;
import com.github.excel.read.handler.ExcelParseHandler;
import com.github.excel.util.StringUtil;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:11 下午
 * @Description: Excel 抽象通用服务
 */
@Slf4j
public abstract class AbstractExcelReader implements ExcelReader {
	/**
	 * 读取excel文件
	 */
	protected InputStream readExcelStream;
	/**
	 * 单个model 缓存
	 */
	protected final Map<Class<? extends ExcelBaseModel>, ExcelReadModelDto> modelMap = Maps.newHashMap();
	/**
	 * modeldto 缓存
	 */
	protected final List<ExcelReadModelDto> readModelDtoList = Lists.newArrayList();
	/**
	 * excel 名称
	 */
	protected final String excelName;
	/**
	 * 最后返回异常
	 */
	protected boolean endReturnException = false;
	/**
	 * 是否读取图片
	 */
	protected boolean readPicture = false;
	/**
	 * 排除器
	 */
	protected String template ;

	public AbstractExcelReader(String excelName, InputStream readExcelStream) {
		if (Objects.isNull(readExcelStream)) {
			throw new ExcelReadException("argument readExcelStream can't be null");
		}
		if (StringUtil.isEmpty(excelName) || (!excelName.endsWith(ExcelConstant.XLSX_STR) && !excelName.endsWith(ExcelConstant.XLS_STR))) {
			throw new ExcelReadException("ExcelName is not correct");
		}
		this.readExcelStream = readExcelStream;
		this.excelName = excelName;
		this.readExcelStream = readExcelStream;
	}

	public AbstractExcelReader(String excelName, InputStream readExcelStream,String template) {
		this(excelName, readExcelStream);
		this.template = template;
	}
	/**
	 * 创建model
	 *
	 * @param modelCla   模型class
	 * @param sheetIndex sheet 下标
	 * @return
	 */
	protected ExcelReadModelDto createModel(Class<? extends ExcelBaseModel> modelCla, int sheetIndex , ExcelReadBatchProcess batchProcess) {
		if (Objects.isNull(modelCla)) {
			throw new ExcelReadException("Model class can't be null");
		}
		if (sheetIndex < ExcelConstant.ZERO_SHORT) {
			throw new ExcelReadException("Sheet index can't be less-than 0");
		}
		ExcelCacheImportModel excelCacheImportModel = ExcelBootLoader.getExcelCacheImportMapValue(modelCla);
		if (Objects.isNull(excelCacheImportModel)) {
			throw new ExcelReadException("Model class unloaded");
		}
		ExcelReadModelDto dto = new ExcelReadModelDto();
		try {
			dto.setModel(modelCla.newInstance());
		} catch (InstantiationException e) {
			log.error("Add model failed , cause :", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException("Add model failed , cause :" + e.getMessage());
		} catch (IllegalAccessException e) {
			log.error("Add model failed , cause :", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException("Add model failed , cause :" + e.getMessage());
		}
		dto.setSheetIndex(sheetIndex);
		dto.setModelCla(modelCla);
		dto.setCacheImportModel(excelCacheImportModel);
		dto.setBatchProcess(batchProcess);
		if (Objects.isNull(batchProcess)) {
			dto.setBatchExecute(false);
		}else{
			dto.setBatchExecute(true);
		}
		return dto;
	}

	@Override
	public void addModel(Class<? extends ExcelBaseModel> modelCla, int sheetIndex) {
		ExcelReadModelDto dto = createModel(modelCla, sheetIndex,null);
		modelMap.put(modelCla, dto);
		readModelDtoList.add(dto);
	}

	@Override
	public void addModelList(Class<? extends ExcelBaseModel> modelCla, int sheetIndex) {
		ExcelReadModelDto dto = createModel(modelCla, sheetIndex,null);
		List<? extends ExcelBaseModel> list = Lists.newArrayList();
		dto.setModelList(list);
		modelMap.put(modelCla, dto);
		readModelDtoList.add(dto);
	}

	@Override
	public void addModelList(Class<? extends ExcelBaseModel> modelCla, int sheetIndex, ExcelReadBatchProcess batchProcess) {
		ExcelReadModelDto dto = createModel(modelCla, sheetIndex,batchProcess);
		List<? extends ExcelBaseModel> list = Lists.newArrayList();
		dto.setModelList(list);
		modelMap.put(modelCla, dto);
		readModelDtoList.add(dto);
	}

	@Override
	public <T extends ExcelBaseModel> T getModel(Class<T> modelCla) {
		ExcelReadModelDto modelDto = modelMap.get(modelCla);
		if (Objects.isNull(modelDto)) {
			return null;
		}
		return (T) modelDto.getModel();
	}

	@Override
	public <T extends ExcelBaseModel> List<T> getModelList(Class<T> modelCla) {
		ExcelReadModelDto modelDto = modelMap.get(modelCla);
		if (Objects.isNull(modelDto)) {
			return null;
		}
		return modelDto.getModelList();
	}

	@Override
	public void setReadPicture(boolean readPicture) {
		this.readPicture = readPicture;
	}

	@Override
	public void parse() {
		ExcelParseHandler handler = this.createHandler();
		handler.process(readModelDtoList, this.endReturnException, this.readPicture,this.template);
	}

	@Override
	public ExcelReadErrorMsgModel parseWithError() {
		this.endReturnException = true;
		ExcelParseHandler handler = this.createHandler();
		return handler.process(readModelDtoList, this.endReturnException,this.readPicture, this.template);
	}


	public abstract ExcelParseHandler createHandler();

	@Data
	public static class ExcelReadModelDto<T extends ExcelBaseModel> {
		private T model;
		private Class<? extends ExcelBaseModel> modelCla;
		private List<T> modelList;
		private Integer sheetIndex;
		private ExcelReadBatchProcess batchProcess;
		private Boolean batchExecute;
		private ExcelCacheImportModel cacheImportModel;
	}
}
