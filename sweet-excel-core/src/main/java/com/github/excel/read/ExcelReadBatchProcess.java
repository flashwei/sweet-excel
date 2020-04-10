package com.github.excel.read;

import com.github.excel.model.ExcelBaseModel;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:13 下午
 * @Description: Excel 读取批处理
 */
public interface ExcelReadBatchProcess<T extends ExcelBaseModel> {
	/**
	 * 获取批次大小
	 *
	 * @return
	 */
	int getBatchSize();

	/**
	 * 批处理
	 *
	 * @param dataList
	 */
	void process(List<T> dataList);

	/**
	 * 执行批处理
	 *
	 * @param dataList
	 */
	default void doProcess(List<T> dataList) {
		if (dataList.size() >= getBatchSize()) {
			process(dataList);
			dataList.clear();
		}
	}
}
