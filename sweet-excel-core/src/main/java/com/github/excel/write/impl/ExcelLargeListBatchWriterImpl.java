package com.github.excel.write.impl;

import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelWriteException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.util.ExcelUtil;
import com.github.excel.util.ZipCompressUtil;
import com.github.excel.write.AbstractExcelStyle;
import com.github.excel.write.BaseExcelWriter;
import com.github.excel.write.ExcelBasicStyle;
import com.github.excel.write.ExcelLargeListBatchWriter;
import com.github.excel.write.ExcelWriterDataFormat;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Future;
import java.util.concurrent.SynchronousQueue;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicLong;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:16 下午
 * @Description: 超大list异步导出多个文件，并压缩
 */
@Slf4j
public class ExcelLargeListBatchWriterImpl extends BaseExcelWriter implements ExcelLargeListBatchWriter {


	private final ExecutorService executorService;

	private String outputDirPath;

	private final List<LargeListAsyncParam> asyncParamList = Lists.newArrayList();

	private final File dirFile;

	private static final String ZIP_SUFFIX = ".zip";

	private final int MAX_POOL_SIZE = 20 ;

	private static final String LARGE_LIST_BATCH_POOL = "ExcelLargeListBatchPool-";

	public ExcelLargeListBatchWriterImpl(String outputDirPath) {
		this(outputDirPath, null);
	}

	public ExcelLargeListBatchWriterImpl(String outputDirPath ,Integer maxPoolSize) {
		if (StringUtil.isEmpty(outputDirPath)) {
			throw new ExcelWriteException("outputDirPath can't be null ");
		}
		if (Objects.isNull(maxPoolSize)) {
			maxPoolSize = MAX_POOL_SIZE;
		}
		dirFile = new File(outputDirPath);
		if (!dirFile.exists()) {
			dirFile.mkdirs();
		}
		this.outputDirPath = outputDirPath;
		this.addStyle(ExcelBasicStyle.class);
		executorService = new ThreadPoolExecutor(ExcelConstant.ONE_INT, maxPoolSize, ExcelConstant.ONE_INT, TimeUnit.MINUTES, new SynchronousQueue<>(), new ThreadFactory() {
			int i = ExcelConstant.ZERO_SHORT;

			@Override
			public Thread newThread(Runnable r) {
				i++;
				return new Thread(r,LARGE_LIST_BATCH_POOL + i);
			}
		});
	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList, String fileName, String sheetName,String[] excludeFields) {
		if (StringUtil.isEmpty(fileName)) {
			throw new ExcelWriteException("fileName can't be null");
		}
		SXSSFWorkbook workbook = new SXSSFWorkbook(ExcelConstant.INT_10000);
		String excelPath = outputDirPath + ExcelConstant.FILE_SEPARATOR + fileName + ExcelConstant.XLSX_STR;
		LargeListAsyncParam param = new LargeListAsyncParam(sheetName, fileName, modelList, workbook, outputDirPath, excelPath,excludeFields);
		LargeListAsyncWriterTask writerTask = new LargeListAsyncWriterTask(param);
		Future<Boolean> future = executorService.submit(writerTask);
		param.setFuture(future);
		asyncParamList.add(param);

	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList, String fileName, String sheetName) {
		process(modelList,fileName,sheetName,null);
	}

	@Override
	public void export(String zipFileName) {
		if (StringUtil.isEmpty(zipFileName)) {
			throw new ExcelWriteException("zipFileName can't be null");
		}
		zipFileName+=ZIP_SUFFIX;
		try {
			List<File> fileList = Lists.newArrayList();
			for (int i = ExcelConstant.ZERO_SHORT; i < asyncParamList.size(); i++) {
				LargeListAsyncParam param = asyncParamList.get(i);
				while (!param.getFuture().isDone()) {
				}
				Boolean result = param.getFuture().get();
				if (Objects.isNull(result) || !result) {
					throw new ExcelWriteException("export excel failed. cause:fill batch no " + i + " error");
				}
				File file = new File(param.getExcelPath());
				fileList.add(file);
			}

			ZipCompressUtil compressUtil = new ZipCompressUtil();
			compressUtil.compressFile(fileList, dirFile + ExcelConstant.FILE_SEPARATOR + zipFileName, null);
			for (File file : fileList) {
				file.delete();
			}
		} catch (InterruptedException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		} catch (ExecutionException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}finally {
			executorService.shutdown();
		}
	}

	@Override
	public void export(HttpServletRequest request, HttpServletResponse response, String zipFileName) {
		try {
			export(zipFileName);
			ExcelUtil.setResponseHeader(request, response, zipFileName, null);
			ServletOutputStream outputStream = response.getOutputStream();
			zipFileName+=ZIP_SUFFIX;
			File zipFile = new File(dirFile + ExcelConstant.FILE_SEPARATOR + zipFileName);
			byte[] zipBytes = Files.readAllBytes(zipFile.toPath());
			outputStream.write(zipBytes);
		} catch (UnsupportedEncodingException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		} catch (IOException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}
	}

	@Override
	public void addStyle(Class<? extends AbstractExcelStyle> styleClass) {
		this.styleList.add(styleClass);
	}

	/**
	 * @Description: 异步导出任务
	 * @Author: Vachel Wang
	 * @Date: 2019/11/5 下午1:21
	 * @Email:
	 */
	private class LargeListAsyncWriterTask implements Callable<Boolean> {
		private LargeListAsyncParam param;

		public LargeListAsyncWriterTask(LargeListAsyncParam param) {
			this.param = param;
		}

		@Override
		public Boolean call() throws Exception {
			SXSSFWorkbook workbook = param.getWorkbook();
			CreationHelper creationHelper = workbook.getCreationHelper();
			// 初始化样式
			initStyle(workbook);
			if (CollectionUtils.isEmpty(param.getModelList())) {
				// 判断是否有数据，没有数据给出默认提示
				addNoResultData(workbook, creationHelper);
				// 保存内容
				try (OutputStream outputStream = new FileOutputStream(new File(param.getExcelPath()))) {
					workbook.write(outputStream);
				}
				return Boolean.TRUE;
			}
			Class<? extends ExcelBaseModel> modelClass = param.getModelList().get(ExcelConstant.ZERO_SHORT).getClass();
			ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(modelClass);

			int rowIndex = ExcelConstant.ZERO_SHORT;
			String sheetName = param.getSheetName();
			String[] excludeFields = param.getExcludeFields();
			Map<String, String> excludeFieldMap = new HashMap<>();
			List<ExcelCacheModel.ExcelCacheFieldModel> cacheFieldModelList = new ArrayList<>();
			if(Objects.nonNull(excludeFields)) {
				for (String excludeField : excludeFields) {
					excludeFieldMap.put(excludeField, ExcelConstant.NULL_STR);
				}
			}
			try {
				// 填充标题
				SXSSFSheet sheet = workbook.createSheet(sheetName);
				SXSSFRow row = sheet.createRow(rowIndex);
				int colIndex = ExcelConstant.ZERO_SHORT;
				AtomicLong incrementSeq = new AtomicLong(ExcelConstant.ZERO_SHORT);
				for (ExcelCacheModel.ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
					ExcelExportCell exportCell = fieldModel.getExportCell();
					if (Objects.nonNull(excludeFieldMap.get(fieldModel.getFieldName()))) {
						continue;
					}
					String titleName = cacheModel.getExcelExport().incrementSequenceTitle();
					if (Objects.nonNull(exportCell)) {
						titleName = exportCell.titleName();
					}
					cacheFieldModelList.add(fieldModel);
					String styleName = fieldModel.getTitleStyleName();
					SXSSFCell cell = row.createCell(colIndex);
					ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
					ExcelHelper.setRowHeight(row,fieldModel.getTitleRowHeight());
					setCellValueAndStyle(cell, titleName, styleName, null, creationHelper);
					++colIndex;
				}
				++rowIndex;
				// 填充内容
				int i = ExcelConstant.ZERO_SHORT ;
				for (ExcelBaseModel model : param.getModelList()) {
					colIndex = ExcelConstant.ZERO_SHORT;
					i++;
					for (ExcelCacheModel.ExcelCacheFieldModel cacheFieldModel : cacheFieldModelList) {
						ExcelExportCell exportCell = cacheFieldModel.getExportCell();
						Object value ;
						short colWidth = ExcelConstant.MINUS_TWO_SHORT;
						String linkName = ExcelConstant.NULL_STR;
						if(Objects.nonNull(exportCell)) {
							colWidth = exportCell.colWidth();
							linkName = exportCell.linkName();
							try {
								value = cacheFieldModel.getGetMethod().invoke(model);
								ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());
								value = formatValue(value, exportCell.formatPattern(), formatter);
							} catch (IllegalAccessException e) {
								log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
								throw new ExcelWriteException(e.getMessage());
							} catch (InvocationTargetException e) {
								log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
								throw new ExcelWriteException(e.getMessage());
							}
						}else{
							value = incrementSeq.incrementAndGet();
						}

						String styleName = i % ExcelConstant.TOW_INT == ExcelConstant.ZERO_SHORT ? cacheFieldModel.getEvenRowStyleName() : cacheFieldModel.getContentStyleName();
						ExcelHelper.setColWidth(sheet, colIndex, colWidth);

						row = (SXSSFRow) ExcelHelper.getRowOrCreate(sheet, rowIndex);
						Cell cell = row.createCell(colIndex);
						ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
						ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
						if (Objects.nonNull(value) && value instanceof byte[]) {
							createPicture(workbook, sheet, cell, (byte[]) value, styleName);
						} else {
							setCellValueAndStyle(cell, value, styleName, linkName, creationHelper);
						}
						colIndex++;
					}
					rowIndex++;
				}
				// 保存内容
				try (OutputStream outputStream = new FileOutputStream(new File(param.getExcelPath()))) {
					workbook.write(outputStream);
				}
			} catch (Exception ex) {
				log.error("export failed thread:{} ,cause:{}", Thread.currentThread().getName(), Throwables.getStackTraceAsString(ex));
				return Boolean.FALSE;
			}finally {
				workbook.close();
				workbook.dispose();
				styleLocal.remove();
				fontLocal.remove();
			}
			return Boolean.TRUE;
		}
	}


	/**
	 * @Description: 异步导出参数
	 * @Author: Vachel Wang
	 * @Date: 2019/11/5 上午11:48
	 * @Email:
	 */
	@Data
	private class LargeListAsyncParam {

		private String sheetName;

		private String fileName;

		private List<? extends ExcelBaseModel> modelList;

		private volatile SXSSFWorkbook workbook;

		private String outputDirPath;

		private String excelPath;

		private Future<Boolean> future;

		private String[] excludeFields;

		public LargeListAsyncParam(String sheetName, String fileName, List<? extends ExcelBaseModel> modelList, SXSSFWorkbook workbook, String outputDirPath, String excelPath, String[] excludeFields) {
			this.sheetName = sheetName;
			this.fileName = fileName;
			this.modelList = modelList;
			this.workbook = workbook;
			this.outputDirPath = outputDirPath;
			this.excelPath = excelPath;
			this.excludeFields = excludeFields;
		}
	}

}
