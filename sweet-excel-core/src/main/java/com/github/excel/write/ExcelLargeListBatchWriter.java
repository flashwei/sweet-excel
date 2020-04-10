package com.github.excel.write;

import com.github.excel.model.ExcelBaseModel;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:17 下午
 * @Description: 导出大数据异步导出
 */
public interface ExcelLargeListBatchWriter {

	/**
	 * 添加导出一个excel
	 *
	 * @param modelList
	 * @param fileName
	 */
	<T extends ExcelBaseModel> void process(List<T> modelList, String fileName,String sheetName,String[] excludeFields);

	/**
	 * 添加导出一个excel
	 *
	 * @param modelList
	 * @param fileName
	 */
	<T extends ExcelBaseModel> void process(List<T> modelList, String fileName,String sheetName);

	/**
	 * 执行导出，导出到文件
	 *
	 * @param zipFileName
	 */
	void export(String zipFileName);

	/**
	 * 执行导出，导出到客户端
	 *
	 * @param request
	 * @param response
	 */
	void export(HttpServletRequest request, HttpServletResponse response, String fileName);

	/**
	 * 添加样式
	 *
	 * @param styleClass
	 */
	void addStyle(Class<? extends AbstractExcelStyle> styleClass);


}
