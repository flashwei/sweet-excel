package com.github.excel.write;

import com.github.excel.model.ExcelBaseModel;
import org.apache.poi.ss.formula.functions.T;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:17 下午
 * @Description: 导出大数据list
 */
public interface ExcelLargeListWriter {

	/**
	 * 添加modelList，list对象类型相同
	 * 请勿传递不同类型的对象
	 * @param modelList
	 * @param <T>
	 */
	<T extends ExcelBaseModel> void  process(List<T> modelList,String[] excludeFields);

	/**
	 * 导出list
	 * @param modelList
	 * @param <T>
	 */
	<T extends ExcelBaseModel> void  process(List<T> modelList);

	/**
	 * 设置没有数据提示
	 * @param noneDataTips
	 */
	void setNoneDataTips(boolean noneDataTips);

	/**
	 * 执行导出，导出到文件
	 * @param outputStream
	 */
	void export(OutputStream outputStream);

	/**
	 * 执行导出，导出到客户端
	 * @param request
	 * @param response
	 */
	void export(HttpServletRequest request, HttpServletResponse response , String fileName);

	/**
	 * 添加样式
	 * @param styleClass
	 */
	void addStyle(Class<? extends AbstractExcelStyle> styleClass);
	/**
	 * 设置sheet名称
	 * @param sheetName sheet名称
	 * @return
	 */
	void setSheetName(String sheetName);

	/**
	 * 关闭
	 */
	void close() ;


}
