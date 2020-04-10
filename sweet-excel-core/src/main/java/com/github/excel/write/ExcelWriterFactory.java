package com.github.excel.write;

import com.github.excel.exception.ExcelWriteException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.util.StringUtil;
import com.github.excel.write.impl.ExcelLargeLisMultiThreadWriterImpl;
import com.github.excel.write.impl.ExcelLargeListBatchWriterImpl;
import com.github.excel.write.impl.ExcelLargeListWriterImpl;
import com.github.excel.write.impl.ExcelUserWriterImpl;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:18 下午
 * @Description: excel导出工厂
 */
public class ExcelWriterFactory {
	/**
	 * 创建用户模式导出
	 *
	 * @return ExcelWriter
	 */
	public static ExcelWriter createUserModelWriter() {
		return new ExcelUserWriterImpl();
	}

	/**
	 * 创建用户模式导出
	 *
	 * @param template 模板
	 * @return ExcelWriter
	 */
	public static ExcelWriter createUserModelWriterWithTemplate(String template) {
		if (StringUtil.isEmpty(template)) {
			throw new ExcelWriteException("Excel template can't be null");
		}
		return new ExcelUserWriterImpl(template);
	}

	/**
	 * 创建大数据模式导出
	 *
	 * @param sheetName sheet名称
	 * @return ExcelWriter
	 */
	public static ExcelLargeListWriter createLargeListWriter(String sheetName) {
		return new ExcelLargeListWriterImpl(sheetName);
	}

	/**
	 * 创建大数据导出并设置sheet最大行数
	 *
	 * @param sheetName
	 * @param sheetRowMaxCount
	 * @return
	 */
	public static ExcelLargeListWriter createLargeListWriter(String sheetName, int sheetRowMaxCount) {
		return new ExcelLargeListWriterImpl(sheetName, sheetRowMaxCount);
	}

	/**
	 * 创建大数据导出并设置sheet最大行数
	 *
	 * @param sheetName sheet 名称
	 * @param sheetRowMaxCount sheet 最大行数
	 * @param listCla 列表class
	 * @return
	 */
	public static ExcelLargeListWriter createLargeListWriter(String sheetName, int sheetRowMaxCount,Class<? extends ExcelBaseModel> listCla) {
		return new ExcelLargeListWriterImpl(sheetName, sheetRowMaxCount,listCla);
	}

	/**
	 * 创建超大数据多线程导出
	 *
	 * @param sheetName
	 * @param maxPoolSize
	 * @return
	 */
	public static ExcelLargeListWriter createLargeListMultiThreadWriter(String sheetName, int maxPoolSize) {
		return new ExcelLargeLisMultiThreadWriterImpl(sheetName, maxPoolSize);
	}

	/**
	 * 创建超大数据多线程导出 并设置sheet最大行数
	 *
	 * @param sheetName        sheet名称
	 * @param maxPoolSize      线程池大小
	 * @param sheetRowMaxCount 单个sheet存放多少行
	 * @return
	 */
	public static ExcelLargeListWriter createLargeListMultiThreadWriter(String sheetName, int maxPoolSize, int sheetRowMaxCount) {
		return new ExcelLargeLisMultiThreadWriterImpl(sheetName, maxPoolSize, sheetRowMaxCount);
	}
	/**
	 * 创建超大数据多线程导出 并设置sheet最大行数
	 *
	 * @param sheetName        sheet名称
	 * @param maxPoolSize      线程池大小
	 * @param sheetRowMaxCount 单个sheet存放多少行
	 * @param listCla 默认列表class
	 * @return
	 */
	public static ExcelLargeListWriter createLargeListMultiThreadWriter(String sheetName, int maxPoolSize, int sheetRowMaxCount,Class<? extends ExcelBaseModel> listCla) {
		return new ExcelLargeLisMultiThreadWriterImpl(sheetName, maxPoolSize, sheetRowMaxCount,listCla);
	}

	/**
	 * 创建大数据模式导出
	 *
	 * @param outputDirPath 导出文件夹名
	 * @return ExcelWriter
	 */
	public static ExcelLargeListBatchWriter createLargeListBatchWriter(String outputDirPath) {
		return new ExcelLargeListBatchWriterImpl(outputDirPath);
	}

	/**
	 * 创建大数据模式导出
	 *
	 * @param outputDirPath 导出文件夹名
	 * @param maxPoolSize   线程池大小
	 * @return ExcelWriter
	 */
	public static ExcelLargeListBatchWriter createLargeListBatchWriter(String outputDirPath, int maxPoolSize) {
		return new ExcelLargeListBatchWriterImpl(outputDirPath, maxPoolSize);
	}

}
