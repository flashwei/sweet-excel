package com.github.excel.boot;

import com.github.excel.write.AbstractExcelStyle;
import com.github.excel.write.ExcelBasicStyle;
import com.google.common.base.Throwables;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelWriteException;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Constructor;
import java.util.HashMap;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 6:59 下午
 * @Description: Workbook缓存池
 */
@Slf4j
public class WorkbookCachePool {

	@Getter
	private static ThreadLocal<WorkbookCacheModel> hssfWorkbookThreadLocal;

	@Getter
	private static ThreadLocal<WorkbookCacheModel> xssfWorkbookThreadLocal;

	@Getter
	private static ThreadLocal<WorkbookCacheModel> sxssfWorkbookThreadLocal;

	/**
	 * 初始化缓存
	 */
	public static void init() {
		hssfWorkbookThreadLocal = addBasicStyle(() -> new HSSFWorkbook());

		xssfWorkbookThreadLocal = addBasicStyle(() -> new XSSFWorkbook());

		sxssfWorkbookThreadLocal = addBasicStyle(() -> new SXSSFWorkbook(ExcelConstant.INT_10000));
	}

	/**
	 * 添加style
	 *
	 * @param workbook workbook
	 * @return
	 */
	public static ThreadLocal<WorkbookCacheModel> addBasicStyle(Workbook workbook) {
		try {
			Class<ExcelBasicStyle> basicStyleClass = ExcelBasicStyle.class;
			Map<String, CellStyle> styleMap = new HashMap<>();
			Map<String, Font> fontMap = new HashMap<>();
			Constructor<? extends AbstractExcelStyle> constructor = basicStyleClass.getConstructor(basicStyleClass.getConstructors()[ExcelConstant.ZERO_SHORT].getParameterTypes());
			AbstractExcelStyle excelStyle = constructor.newInstance(workbook, styleMap, fontMap);
			excelStyle.addNewFont();
			excelStyle.addNewStyle();
			WorkbookCacheModel workbookCacheModel = new WorkbookCacheModel(workbook, styleMap, fontMap);
			return ThreadLocal.withInitial(() -> {
				return workbookCacheModel;
			});
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException("Init style error");
		}
	}

	/**
	 * 添加style
	 *
	 * @return
	 */
	public static ThreadLocal<WorkbookCacheModel> addBasicStyle(WorkbookResolver workbookResolver) {

		Class<ExcelBasicStyle> basicStyleClass = ExcelBasicStyle.class;

		return ThreadLocal.withInitial(() -> {
			Workbook workbook = workbookResolver.resolve();
			try {
				Map<String, CellStyle> styleMap = new HashMap<>();
				Map<String, Font> fontMap = new HashMap<>();
				Constructor<? extends AbstractExcelStyle> constructor = basicStyleClass.getConstructor(basicStyleClass.getConstructors()[ExcelConstant.ZERO_SHORT].getParameterTypes());
				AbstractExcelStyle excelStyle = constructor.newInstance(workbook, styleMap, fontMap);
				excelStyle.addNewFont();
				excelStyle.addNewStyle();
				WorkbookCacheModel workbookCacheModel = new WorkbookCacheModel(workbook, styleMap, fontMap);
				return workbookCacheModel;
			} catch (Exception e) {
				log.error(Throwables.getStackTraceAsString(e));
				throw new ExcelWriteException("Init style error");
			}
		});

	}

	/**
	 * @Author: Vachel Wang
	 * @Date: 2020/4/2 6:59 下午
	 * @Description: workbook 缓存模型
	 */
	@AllArgsConstructor
	public static class WorkbookCacheModel {
		@Getter
		private Workbook workbook;
		@Getter
		private Map<String, CellStyle> styleMap;
		@Getter
		private Map<String, Font> fontMap;
	}

	@FunctionalInterface
	public interface WorkbookResolver{
		Workbook resolve();
	}
}
