package com.github.excel.write;

import com.github.excel.model.ExcelCustomColumnModel;
import com.github.excel.model.ExcelMergeCustomColumnModel;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.model.ExcelBaseModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:17 下午
 * @Description: 导出接口
 */
public interface ExcelWriter {
	/**
	 * 添加单个model
	 * @param model
	 * @param <T>
	 */
	<T extends ExcelBaseModel> BaseExcelWriter.ExcludeHandler addModel(T model, String sheetName, boolean fillTemplate);
	/**
	 * 添加单个model
	 * @param model
	 * @param <T>
	 * @param rowIndex
	 * @param colIndex
	 */
	<T extends ExcelBaseModel> BaseExcelWriter.ExcludeHandler  addModel(int rowIndex, int colIndex, T model, String sheetName, boolean fillTemplate);

	/**
	 * 添加modelList，list对象类型相同
	 * 请勿传递不同类型的对象
	 * @param modelList
	 * @param <T>
	 */
	<T extends ExcelBaseModel> BaseExcelWriter.ExcludeHandler  addModelList(List<T> modelList, String sheetName, boolean fillTemplate);

	/**
	 * 添加modelList，list对象类型相同
	 * 请勿传递不同类型的对象
	 * @param modelList
	 * @param <T>
	 * @param rowIndex
	 * @param colIndex
	 */
	<T extends ExcelBaseModel> BaseExcelWriter.ExcludeHandler  addModelList(int rowIndex, int colIndex, List<T> modelList, String sheetName, boolean fillTemplate);

	/**
	 * 设置没有数据提示
	 * @param noneDataTips
	 */
	void setNoneDataTips(boolean noneDataTips);

	/**
	 * 添加自定义单元格
	 * @param customColumnModel
	 */
	void addCustomColumn(ExcelCustomColumnModel customColumnModel);

	/**
	 * 合并单元格并设值
	 * @param mergeCustomColumnModel
	 */
	void addMergeCustomColumn(ExcelMergeCustomColumnModel mergeCustomColumnModel);

	/**
	 * 执行导出，导出到文件
	 * @param outputStream
	 */
	void process(OutputStream outputStream, String fileName, ExcelSuffixEnum suffixEnum);

	/**
	 * 执行导出，导出到客户端
	 * @param request
	 * @param response
	 * @param fileName
	 * @param suffixEnum
	 */
	void process(HttpServletRequest request, HttpServletResponse response, String fileName, ExcelSuffixEnum suffixEnum);

	/**
	 * excel 自定义写
	 * @param customWrite
	 */
	void setCustomWrite(ExcelCustomWriter customWrite);

	/**
	 * 添加样式
	 * @param styleClass
	 */
	void addStyle(Class<? extends AbstractExcelStyle> styleClass);

	/**
	 * 获取样式
	 * @param name 样式名称
	 * @return
	 */
	CellStyle getStyle(String name);

	/**
	 * 获取字体
	 * @param name 字体名称
	 * @return
	 */
	Font getFont(String name);

	/**
	 * 大数据量简单导出  默认false
	 * 	 只用于基于海量数据的注解导出模式
	 * @param streaming
	 */
	void setStreaming(boolean streaming);

	/**
	 * 选中sheet
	 * @param sheetName sheetName
	 * @return 是否选中成功
	 */
	void selectSheet(String sheetName);

	/**
	 * 列表class
	 * @param listCla
	 */
	void setListCla(Class<? extends ExcelBaseModel> listCla);
}
