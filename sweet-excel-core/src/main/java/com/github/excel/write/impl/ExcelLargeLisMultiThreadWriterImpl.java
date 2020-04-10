package com.github.excel.write.impl;

import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.boot.WorkbookCachePool;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.util.ExcelUtil;
import com.github.excel.write.AbstractExcelStyle;
import com.github.excel.write.BaseExcelWriter;
import com.github.excel.write.ExcelWriterDataFormat;
import com.google.common.base.Throwables;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelWriteException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelLargeListWriter;
import com.google.common.collect.Maps;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.Callable;
import java.util.concurrent.CountDownLatch;
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
 * @Description: excel 导出大list多线程实现.
 */
@Slf4j
public class ExcelLargeLisMultiThreadWriterImpl extends BaseExcelWriter implements ExcelLargeListWriter {

	private final SXSSFWorkbook workbook;

	private final ThreadLocal<WorkbookCachePool.WorkbookCacheModel> sxssfWorkbookThreadLocal ;

	private volatile String sheetName;

	private boolean fillTitle = true;

	private volatile int sheetNo = ExcelConstant.ONE_INT;

	private volatile int rowIndex = ExcelConstant.ZERO_SHORT;

	private ExecutorService executorService ;

	private List<Future<Boolean>> futureList = new ArrayList<>();

	private CountDownLatch countDownLatch ;

	private int sheetRowMaxCount;

	private int totalNum ;

	private static final String LARGE_LIST_MULTI_THREAD_POOL = "ExcelLargeListMultiPool-";

	public ExcelLargeLisMultiThreadWriterImpl(String sheetName , int totalNum) {
		sxssfWorkbookThreadLocal = WorkbookCachePool.getSxssfWorkbookThreadLocal();
		WorkbookCachePool.WorkbookCacheModel workbookCacheModel = sxssfWorkbookThreadLocal.get();
		workbook = (SXSSFWorkbook) workbookCacheModel.getWorkbook();
		this.styleLocal.set(workbookCacheModel.getStyleMap());
		this.fontLocal.set(workbookCacheModel.getFontMap());

		this.sheetName = sheetName;
		this.totalNum = totalNum;
		if (totalNum < ExcelConstant.ONE_INT) {
			return ;
		}
		executorService = new ThreadPoolExecutor(ExcelConstant.ONE_INT, totalNum, ExcelConstant.ONE_INT, TimeUnit.MINUTES, new SynchronousQueue<>(), new ThreadFactory() {
			int i = ExcelConstant.ZERO_SHORT;

			@Override
			public Thread newThread(Runnable r) {
				i++;
				return new Thread(r,LARGE_LIST_MULTI_THREAD_POOL + i);
			}
		});
		countDownLatch = new CountDownLatch(totalNum);
	}

	public ExcelLargeLisMultiThreadWriterImpl(String sheetName , int totalNum,int sheetRowMaxCount) {
		this(sheetName, totalNum);
		if (sheetRowMaxCount > ExcelConstant.ZERO_SHORT && sheetRowMaxCount <= ExcelConstant.INT_1000000) {
			this.sheetRowMaxCount = sheetRowMaxCount;
		}else{
			this.sheetRowMaxCount = ExcelConstant.INT_1000000 ;
		}
	}
	public ExcelLargeLisMultiThreadWriterImpl(String sheetName , int totalNum,int sheetRowMaxCount,Class<? extends ExcelBaseModel> listCla) {
		this(sheetName, totalNum,sheetRowMaxCount);
		this.listCla = listCla;
	}

	/**
	 * 此方法只能同步调用
	 * @param modelList
	 * @param excludeFields
	 * @param <T>
	 */
	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList,String[] excludeFields) {
		if (totalNum < ExcelConstant.ONE_INT) {
			return ;
		}

		if (Objects.isNull(modelList) || modelList.isEmpty()) {
			throw new ExcelWriteException("modelList can't be null");
		}

		SXSSFSheet sheet ;
		String currentSheetName = sheetName + ExcelConstant.SHORT_TERM + sheetNo;;
		if (fillTitle) {
			rowIndex = ExcelConstant.ZERO_SHORT;

			sheet = workbook.createSheet(currentSheetName);
		}else{
			sheet = workbook.getSheet(currentSheetName);
		}

		Map<String, CellStyle> styleMap = styleLocal.get();
		Map<String, Font> fontMap = fontLocal.get();


		Future<Boolean> future = executorService.submit(new LargeListWriterTask(modelList, excludeFields, sheet, rowIndex,countDownLatch,fillTitle,styleMap,fontMap));
		futureList.add(future);

		if (fillTitle) {
			rowIndex ++ ;
			fillTitle = false;
		}
		rowIndex += modelList.size();

		if (rowIndex >= sheetRowMaxCount) {
			sheetNo++;
			fillTitle = true;
		}
	}

	private class LargeListWriterTask implements Callable<Boolean> {
		private List<? extends ExcelBaseModel> modelList ;
		private String[] excludeFields ;
		private SXSSFSheet sheet ;
		private int startRowIndex ;
		private CountDownLatch countDownLatch ;
		private boolean isFillTitle ;
		private Map<String, CellStyle> styleMap;
		private Map<String, Font> fontMap;
		private Logger LOG = LoggerFactory.getLogger(LargeListWriterTask.class);
		public LargeListWriterTask(List<? extends ExcelBaseModel> modelList, String[] excludeFields, SXSSFSheet sheet, int startRowIndex, CountDownLatch countDownLatch, boolean isFillTitle, Map<String, CellStyle> styleMap, Map<String, Font> fontMap) {
			this.modelList = modelList ;
			this.excludeFields = excludeFields;
			this.sheet = sheet ;
			this.startRowIndex = startRowIndex ;
			this.countDownLatch = countDownLatch ;
			this.isFillTitle = isFillTitle;
			this.styleMap = styleMap;
			this.fontMap = fontMap;
		}

		@Override
		public Boolean call() {
			try {
				// 设置样式&字体
				styleLocal.set(styleMap);
				fontLocal.set(fontMap);
				Map<String, String> excludeFieldMap = new HashMap<>();
				if(Objects.nonNull(excludeFields)) {
					for (String excludeField : excludeFields) {
						excludeFieldMap.put(excludeField, ExcelConstant.NULL_STR);
					}
				}
				Class<? extends ExcelBaseModel> modelClass = modelList.get(ExcelConstant.ZERO_SHORT).getClass();
				ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(modelClass);
				AtomicLong seqNo = new AtomicLong(ExcelConstant.ZERO_SHORT);

				CreationHelper creationHelper = workbook.getCreationHelper();
				List<ExcelCacheModel.ExcelCacheFieldModel> cacheFieldModelList = new ArrayList<>();
				for (ExcelCacheModel.ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
					if (Objects.nonNull(excludeFieldMap.get(fieldModel.getFieldName()))) {
						continue;
					}
					cacheFieldModelList.add(fieldModel);
				}
				// 填充标题
				if (isFillTitle) {
					SXSSFRow row = this.sheet.createRow(startRowIndex);
					int colIndex = ExcelConstant.ZERO_SHORT;
					for (ExcelCacheModel.ExcelCacheFieldModel fieldModel : cacheFieldModelList) {
						ExcelExportCell exportCell = fieldModel.getExportCell();
						if (Objects.nonNull(excludeFieldMap.get(fieldModel.getFieldName()))) {
							continue;
						}
						String titleName = cacheModel.getExcelExport().incrementSequenceTitle();
						if (Objects.nonNull(exportCell)) {
							titleName = exportCell.titleName();
						}
						String styleName = fieldModel.getTitleStyleName();
						SXSSFCell cell = row.createCell(colIndex);
						ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
						ExcelHelper.setRowHeight(row,fieldModel.getTitleRowHeight());
						setCellValueAndStyle(cell, titleName, styleName, null, creationHelper);
//						cell.setCellValue(fieldModel.getExportCell().titleName());
						++colIndex;
					}
					++startRowIndex;
				}
				// 填充内容
				int i = ExcelConstant.ZERO_SHORT;
				for (ExcelBaseModel model : modelList) {
					int colIndex = ExcelConstant.ZERO_SHORT;
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
							value = seqNo.incrementAndGet();
						}

						String styleName = i % ExcelConstant.TOW_INT == ExcelConstant.ZERO_SHORT ? cacheFieldModel.getEvenRowStyleName() : cacheFieldModel.getContentStyleName();
						ExcelHelper.setColWidth(this.sheet, colIndex,colWidth);

						Row row = ExcelHelper.getRowOrCreate(this.sheet, startRowIndex);
						Cell cell = row.createCell(colIndex);
						ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
						ExcelHelper.setColWidth(this.sheet, colIndex, cacheFieldModel.getColWidth());
						if (Objects.nonNull(value) && value instanceof byte[]) {
							createPicture(workbook, this.sheet, cell, (byte[]) value, styleName);
						} else {
							setCellValueAndStyle(cell, value, styleName, linkName, creationHelper);
						}
						colIndex++;
					}
					colIndex = ExcelConstant.ZERO_SHORT;
					startRowIndex++;
				}
			} catch (Throwable e) {
				LOG.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
				return false ;
			}finally {
				countDownLatch.countDown();
			}
			return true;
		}
	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList) {
		process(modelList,null);
	}

    @Override
    public void setNoneDataTips(boolean noneDataTips) {
		this.noneDataTips = noneDataTips;
    }

    @Override
	public void export(OutputStream outputStream) {
		boolean result = getResult();
		if (!result) {
			throw new ExcelWriteException("export failed");
		}
		CreationHelper creationHelper = workbook.getCreationHelper();
		addNoResultData(workbook, creationHelper);
		try {
			workbook.write(outputStream);
		} catch (IOException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}finally {
			try {
				if(Objects.nonNull(executorService)) {
					executorService.shutdown();
				}

				outputStream.flush();
				outputStream.close();
				workbook.dispose();
				styleLocal.remove();
				fontLocal.remove();
				sxssfWorkbookThreadLocal.remove();
			} catch (IOException e) {
				log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
				throw new ExcelWriteException(e.getMessage());
			}
		}
	}


	@Override
	public void export(HttpServletRequest request, HttpServletResponse response, String fileName) {
		boolean result = getResult();
		if (!result) {
			throw new ExcelWriteException("export failed");
		}
		CreationHelper creationHelper = workbook.getCreationHelper();
		addNoResultData(workbook, creationHelper);
		try {
			ExcelUtil.setResponseHeader(request, response, fileName, ExcelConstant.XLSX_STR);
			OutputStream outputStream = response.getOutputStream();
			workbook.write(outputStream);
		} catch (IOException e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}finally {
			if(Objects.nonNull(executorService)) {
				executorService.shutdown();
			}
			workbook.dispose();
			fontLocal.remove();
			styleLocal.remove();
			sxssfWorkbookThreadLocal.remove();
		}
	}

	private boolean getResult() {
		if (totalNum < ExcelConstant.ONE_INT) {
			return true;
		}
		boolean result = true ;
		try {
			countDownLatch.await(ExcelConstant.INT_30,TimeUnit.SECONDS);
		} catch (InterruptedException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriteException(e.getMessage());
		}
		for (int i = ExcelConstant.ZERO_SHORT; i<futureList.size(); i++) {
			Future<Boolean> future = futureList.get(i);
			try {
				boolean executeResult =  future.get();
				if (!executeResult) {
					result = executeResult;
				}
				log.info("export {} by batchNo {}",executeResult ==false ? "failed":"success",i+ExcelConstant.ONE_INT);
			} catch (InterruptedException e) {
				log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
				throw new ExcelWriteException(e.getMessage());
			} catch (ExecutionException e) {
				log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
				throw new ExcelWriteException(e.getMessage());
			}
		}
		return result;
	}

	@Override
	public void addStyle(Class<? extends AbstractExcelStyle> styleClass) {
		initStyle(workbook,styleClass);
	}

	@Override
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	@Override
	public void close() {

	}
}
