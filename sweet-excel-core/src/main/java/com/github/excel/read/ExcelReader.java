package com.github.excel.read;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReadErrorMsgModel;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:13 下午
 * @Description: Excel 读取
 */
public interface ExcelReader {
	/**
	 * 添加单个bean
	 *
	 * @param modelCla
	 * @param sheetIndex
	 */
	void addModel(Class<? extends ExcelBaseModel> modelCla, int sheetIndex);

	/**
	 * 添加bean list
	 *
	 * @param modelCla
	 * @param sheetIndex
	 */
	void addModelList(Class<? extends ExcelBaseModel> modelCla, int sheetIndex);

	/**
	 * 添加bean list，并设置批处理
	 *
	 * @param modelCla
	 * @param sheetIndex
	 * @param batchProcess
	 */
	void addModelList(Class<? extends ExcelBaseModel> modelCla, int sheetIndex, ExcelReadBatchProcess batchProcess);

	/**
	 * 获取单个bean
	 *
	 * @param modelCla
	 * @param <T>
	 * @return
	 */
	<T extends ExcelBaseModel> T getModel(Class<T> modelCla);

	/**
	 * 获取bean list
	 *
	 * @param modelCla
	 * @param <T>
	 * @return
	 */
	<T extends ExcelBaseModel> List<T> getModelList(Class<T> modelCla);

	/**
	 * 设置是否支持读取图片
	 *
	 * @param readPicture
	 */
	void setReadPicture(boolean readPicture);

	/**
	 * 执行解析，遇到错误中断
	 */
	void parse();

	/**
	 * 执行解析，遇到错误统计，最后统一返回
	 *
	 * @return
	 */
	ExcelReadErrorMsgModel parseWithError();
}
