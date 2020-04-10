package com.github.excel.write.impl;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.exception.ExcelWriteException;
import com.github.excel.model.ExcelCustomColumnModel;
import com.github.excel.model.ExcelMergeCustomColumnModel;
import com.github.excel.util.ExcelUtil;
import com.github.excel.util.StringUtil;
import com.github.excel.write.AbstractExcelStyle;
import com.github.excel.write.BaseExcelWriter;
import com.github.excel.write.ExcelCustomWriter;
import com.github.excel.write.ExcelWriter;
import com.google.common.base.Throwables;
import com.github.excel.model.ExcelBaseModel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:16 下午
 * @Description: 用户模式下导出excel
 */
@Slf4j
public class ExcelUserWriterImpl extends BaseExcelWriter implements ExcelWriter {

	public ExcelUserWriterImpl(String template) {
		this.template = template;
	}

	public ExcelUserWriterImpl() {

	}

	/**
	 * 添加导出模型
	 *
	 * @param model     模型
	 * @param sheetName sheet名称
	 */
	@Override
	public <T extends ExcelBaseModel> ExcludeHandler addModel(T model, String sheetName, boolean fillTemplate) {
		if (Objects.isNull(model)) {
			throw new ExcelWriteException("model can not be null!");
		}
		addModel(model.getClass(), sheetName, model, null, null, null, fillTemplate, exportBeanList);
		return new ExcludeHandler(model.getClass());
	}

	/**
	 * 添加导出模型
	 *
	 * @param rowIndex  行号
	 * @param colIndex  列号
	 * @param model     模型
	 * @param sheetName sheet名称
	 */
	@Override
	public <T extends ExcelBaseModel> ExcludeHandler addModel(int rowIndex, int colIndex, T model, String sheetName, boolean fillTemplate) {
		if (Objects.isNull(model)) {
			throw new ExcelWriteException("model can not be null!");
		}
		addModel(model.getClass(), sheetName, model, null, rowIndex, colIndex, fillTemplate, exportBeanList);
		return new ExcludeHandler(model.getClass());
	}

	/**
	 * 添加导出模型List
	 *
	 * @param modelList 模型list
	 * @param sheetName sheet名称
	 */
	@Override
	public <T extends ExcelBaseModel> ExcludeHandler addModelList(List<T> modelList, String sheetName, boolean fillTemplate) {
		if (Objects.isNull(modelList) || modelList.size() == ExcelConstant.ZERO_SHORT) {
			throw new ExcelWriteException("model can not be null!");
		}
		Class<? extends ExcelBaseModel> modelCla = modelList.get(ExcelConstant.ZERO_SHORT).getClass();
		addModel(modelCla, sheetName, null, modelList, null, null, fillTemplate, exportModelList);
		return new ExcludeHandler(modelCla);
	}

	/**
	 * 添加导出模型List
	 *
	 * @param rowIndex  行号
	 * @param colIndex  列号
	 * @param modelList list
	 * @param sheetName sheet名称
	 */
	@Override
	public <T extends ExcelBaseModel> ExcludeHandler addModelList(int rowIndex, int colIndex, List<T> modelList, String sheetName, boolean fillTemplate) {
		if (Objects.isNull(modelList) || modelList.size() == ExcelConstant.ZERO_SHORT) {
			throw new ExcelWriteException("model can not be null!");
		}
		Class<? extends ExcelBaseModel> modelCla = modelList.get(ExcelConstant.ZERO_SHORT).getClass();
		addModel(modelCla, sheetName, null, modelList, rowIndex, colIndex, fillTemplate, exportModelList);
		return new ExcludeHandler(modelCla);
	}

	@Override
	public void setNoneDataTips(boolean noneDataTips) {
		super.noneDataTips = noneDataTips;
	}

	/**
	 * 添加自定义列
	 *
	 * @param customColumnModel 自定义列对象
	 */
	@Override
	public void addCustomColumn(ExcelCustomColumnModel customColumnModel) {
		if (StringUtil.isEmpty(customColumnModel.getSheetName())) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_SHEET_NAME);
		}
		if (ExcelConstant.ZERO_SHORT > customColumnModel.getColIndex() || ExcelConstant.ZERO_SHORT > customColumnModel.getRowIndex()) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_COLUMN_POINT);
		}
		if (Objects.isNull(customColumnModel.getValue())) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_COLUMN_VALUE_NULL);
		}
		customColumnModelList.add(customColumnModel);
	}

	/**
	 * 添加自定义合并列
	 *
	 * @param mergeCustomColumnModel
	 */
	@Override
	public void addMergeCustomColumn(ExcelMergeCustomColumnModel mergeCustomColumnModel) {
		if (StringUtil.isEmpty(mergeCustomColumnModel.getSheetName())) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_SHEET_NAME);
		}
		if (ExcelConstant.ZERO_SHORT > mergeCustomColumnModel.getFirstRow() || ExcelConstant.ZERO_SHORT > mergeCustomColumnModel.getLastRow()) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_COLUMN_POINT);
		}
		if (ExcelConstant.ZERO_SHORT > mergeCustomColumnModel.getFirstColumn() || ExcelConstant.ZERO_SHORT > mergeCustomColumnModel.getLastColumn()) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_COLUMN_POINT);
		}
		if (Objects.isNull(mergeCustomColumnModel.getValue())) {
			throw new ExcelWriteException(ExcelErrorMsgConstant.ERROR_COLUMN_VALUE_NULL);
		}
		this.mergeCustomColumnModelList.add(mergeCustomColumnModel);
	}

	/**
	 * 执行导出操作
	 *
	 * @param outputStream 导出流
	 */
	@Override
	public void process(OutputStream outputStream, String fileName, ExcelSuffixEnum suffixEnum) {
		String excelName = fileName + suffixEnum.getSuffix();
		if (StringUtil.notEmpty(template)) {
			writeToTemplate(outputStream);
		} else {
			writeToNewFile(outputStream, excelName);
		}
	}

	@Override
	public void process(HttpServletRequest request, HttpServletResponse response, String fileName, ExcelSuffixEnum suffixEnum) {
		try {
			ExcelUtil.setResponseHeader(request, response, fileName, suffixEnum.getSuffix());
			OutputStream outputStream = response.getOutputStream();
			String excelName = fileName + suffixEnum.getSuffix();
			if (StringUtil.notEmpty(template)) {
				writeToTemplate(outputStream);
			} else {
				writeToNewFile(outputStream, excelName);
			}
		} catch (IOException e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}
	}

	/**
	 * 添加样式
	 *
	 * @param styleClass 样式class
	 */
	@Override
	public void addStyle(Class<? extends AbstractExcelStyle> styleClass) {
		this.styleList.add(styleClass);
	}

	@Override
	public CellStyle getStyle(String name) {
		return this.styleLocal.get().get(name);
	}

	@Override
	public Font getFont(String name) {
		return this.fontLocal.get().get(name);
	}

	@Override
	public void setStreaming(boolean streaming) {
		super.streaming = streaming;
	}

	@Override
	public void selectSheet(String sheetName) {
		this.selectSheet = sheetName;
	}

	@Override
	public void setListCla(Class<? extends ExcelBaseModel> listCla) {
		this.listCla = listCla;
	}

	/**
	 * 设置自定义写器
	 *
	 * @param customWrite 自定义写器
	 */
	@Override
	public void setCustomWrite(ExcelCustomWriter customWrite) {
		this.customWrite = customWrite;
	}


}
