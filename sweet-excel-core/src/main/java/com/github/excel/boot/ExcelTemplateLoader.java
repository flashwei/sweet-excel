package com.github.excel.boot;

import com.github.excel.model.ExcelTemplateTitleModel;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.model.ExcelExpressionModel;
import com.github.excel.model.ExcelImportTemplateCacheModel;
import com.github.excel.model.ExcelTemplateCacheModel;
import com.github.excel.read.AbstractImportTemplateExcluder;
import com.github.excel.util.StringUtil;
import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.Callable;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Future;
import java.util.concurrent.SynchronousQueue;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 6:58 下午
 * @Description: excel模板加载器
 */
@Slf4j
public class ExcelTemplateLoader {

	private static final String LOAD_EXPORT_TEMPLATE_POOL_PREFIX = "excel-export-template-pool-";
	private static final String LOAD_IMPORT_TEMPLATE_POOL_PREFIX = "excel-import-template-pool-";


	/**
	 * 加载导出模板
	 *
	 * @param templates                 模板文件
	 * @param templateCacheMap          模板表达式缓存map
	 * @param templateTitleCacheMap     模板标题缓存map
	 * @param excelTemplateFileCacheMap 模板文件缓存map
	 */
	public static void loadExportTemplate(List<File> templates, Map<String, Map<String, List<ExcelExpressionModel>>> templateCacheMap, Map<String, Map<Integer, List<ExcelTemplateTitleModel>>> templateTitleCacheMap, Map<String, ExcelTemplateCacheModel> excelTemplateFileCacheMap) {
		if (templates.isEmpty()) {
			return;
		}
		List<TaskExcuteResult> futureList = new ArrayList<>(templates.size());
		Map<String, List<ExcelExpressionModel>> excelExpressionMap = Maps.newHashMap();
		Map<String, List<ExcelTemplateTitleModel>> excelTitleMap = Maps.newHashMap();
		ExecutorService executorService = null;
		try {
			executorService = new ThreadPoolExecutor(templates.size(), templates.size(), ExcelConstant.INT_1000, TimeUnit.MILLISECONDS, new SynchronousQueue<>(), new ThreadFactory() {
				int i = ExcelConstant.ZERO_SHORT;

				@Override
				public Thread newThread(Runnable r) {
					i++;
					return new Thread(r, LOAD_EXPORT_TEMPLATE_POOL_PREFIX + i);
				}
			});
			for (File excelFile : templates) {
				List<ExcelExpressionModel> expressionModelList = Lists.newArrayList();
				List<ExcelTemplateTitleModel> titleModelList = Lists.newArrayList();
				excelExpressionMap.put(excelFile.getName(), expressionModelList);
				excelTitleMap.put(excelFile.getName(), titleModelList);
				Future<Boolean> future = executorService.submit(new LoadExportTemplateTask(new FileInputStream(excelFile) , new FileInputStream(excelFile), excelFile.getName(), expressionModelList, titleModelList, excelTemplateFileCacheMap));
				TaskExcuteResult excuteResult = TaskExcuteResult.builder().result(future).templateName(excelFile.getName()).build();
				futureList.add(excuteResult);
			}
			for (TaskExcuteResult excuteResult : futureList) {
				while (!excuteResult.result.isDone()) {
				}
				log.info("Load excel export template {} by {} ", excuteResult.result.get() ? "successfully" : "failed", excuteResult.templateName);
			}
		} catch (Exception ex) {
			log.error("Load excel export template failed,cause:{}", Throwables.getStackTraceAsString(ex));
		} finally {
			if (Objects.nonNull(executorService)) {
				executorService.shutdown();
			}
		}
		excelExpressionMap.forEach((k, v) -> {
			Map<String, List<ExcelExpressionModel>> expressionMap = v.stream().collect(Collectors.groupingBy(ExcelExpressionModel::getNameSpace));
			templateCacheMap.put(k, expressionMap);
		});
		excelTitleMap.forEach((k, v) -> {
			Map<Integer, List<ExcelTemplateTitleModel>> titleMap = v.stream().collect(Collectors.groupingBy(ExcelTemplateTitleModel::getRowIndex));
			templateTitleCacheMap.put(k, titleMap);
		});

	}

	/**
	 * 加载导入模板
	 *
	 * @param templates                       模板文件list
	 * @param excelImportTemplateFileCacheMap 模板文件缓存map
	 * @param excelImportTemplateCacheMap     模板缓存map
	 * @param excluder                        模板缓存排除器
	 */
	public static void loadImportTemplate(List<File> templates, Map<String, byte[]> excelImportTemplateFileCacheMap, Map<String, Map<String, List<ExcelImportTemplateCacheModel>>> excelImportTemplateCacheMap, AbstractImportTemplateExcluder excluder) {
		if (templates.isEmpty()) {
			return;
		}
		ExecutorService executorService = new ThreadPoolExecutor(templates.size(), templates.size(), ExcelConstant.INT_1000, TimeUnit.MILLISECONDS, new SynchronousQueue<>(), new ThreadFactory() {
			int i = ExcelConstant.ZERO_SHORT;

			@Override
			public Thread newThread(Runnable r) {
				i++;
				return new Thread(r, LOAD_IMPORT_TEMPLATE_POOL_PREFIX + i);
			}
		});
		CountDownLatch countDownLatch = new CountDownLatch(templates.size());
		List<TaskExcuteResult> futureList = new ArrayList<>(templates.size());
		try {
			for (File template : templates) {
				Future<Boolean> future = executorService.submit(new LoadImportTemplateTask(new FileInputStream(template) , new FileInputStream(template), template.getName(), excelImportTemplateFileCacheMap, countDownLatch, excelImportTemplateCacheMap, excluder));
				TaskExcuteResult excuteResult = TaskExcuteResult.builder().result(future).templateName(template.getName()).build();
				futureList.add(excuteResult);
			}
			countDownLatch.await();
			for (TaskExcuteResult excuteResult : futureList) {
				log.info("Load excel import template {} by {} ", excuteResult.result.get() ? "successfully" : "failed", excuteResult.templateName);
			}
		} catch (FileNotFoundException e){
			log.error("Load excel import template failed,cause:{}", Throwables.getStackTraceAsString(e));
		} catch (InterruptedException e) {
			log.error("Load excel import template failed,cause:{}", Throwables.getStackTraceAsString(e));
		} catch (ExecutionException e) {
			log.error("Load excel import template failed,cause:{}", Throwables.getStackTraceAsString(e));
		} finally {
			if (Objects.nonNull(executorService)) {
				executorService.shutdown();
			}
		}
	}

	/**
	 * 加载导出模板
	 *
	 * @param templates                 模板文件
	 * @param templateCacheMap          模板表达式缓存map
	 * @param templateTitleCacheMap     模板标题缓存map
	 * @param excelTemplateFileCacheMap 模板文件缓存map
	 */
	public static void loadBootExportTemplate(List<Map<String, Object>> templates, Map<String, Map<String, List<ExcelExpressionModel>>> templateCacheMap, Map<String, Map<Integer, List<ExcelTemplateTitleModel>>> templateTitleCacheMap, Map<String, ExcelTemplateCacheModel> excelTemplateFileCacheMap) {
		if (templates.isEmpty()) {
			return;
		}
		List<TaskExcuteResult> futureList = new ArrayList<>(templates.size());
		Map<String, List<ExcelExpressionModel>> excelExpressionMap = Maps.newHashMap();
		Map<String, List<ExcelTemplateTitleModel>> excelTitleMap = Maps.newHashMap();
		ExecutorService executorService = null;
		try {
			executorService = new ThreadPoolExecutor(templates.size(), templates.size(), ExcelConstant.INT_1000, TimeUnit.MILLISECONDS, new SynchronousQueue<>(), new ThreadFactory() {
				int i = ExcelConstant.ZERO_SHORT;

				@Override
				public Thread newThread(Runnable r) {
					i++;
					return new Thread(r, LOAD_EXPORT_TEMPLATE_POOL_PREFIX + i);
				}
			});
			for (Map<String, Object> template : templates) {
				String excelName = (String)template.get("name");
				InputStream inputStream = (InputStream)template.get("input");
				InputStream cacheInput = (InputStream)template.get("cacheInput");
				List<ExcelExpressionModel> expressionModelList = Lists.newArrayList();
				List<ExcelTemplateTitleModel> titleModelList = Lists.newArrayList();
				excelExpressionMap.put(excelName, expressionModelList);
				excelTitleMap.put(excelName, titleModelList);
				Future<Boolean> future = executorService.submit(new LoadExportTemplateTask(inputStream, cacheInput, excelName, expressionModelList, titleModelList, excelTemplateFileCacheMap));
				TaskExcuteResult excuteResult = TaskExcuteResult.builder().result(future).templateName(excelName).build();
				futureList.add(excuteResult);
			}
			for (TaskExcuteResult excuteResult : futureList) {
				while (!excuteResult.result.isDone()) {
				}
				log.info("Load excel export template {} by {} ", excuteResult.result.get() ? "successfully" : "failed", excuteResult.templateName);
			}
		} catch (Exception ex) {
			log.error("Load excel export template failed,cause:{}", Throwables.getStackTraceAsString(ex));
		} finally {
			if (Objects.nonNull(executorService)) {
				executorService.shutdown();
			}
		}
		excelExpressionMap.forEach((k, v) -> {
			Map<String, List<ExcelExpressionModel>> expressionMap = v.stream().collect(Collectors.groupingBy(ExcelExpressionModel::getNameSpace));
			templateCacheMap.put(k, expressionMap);
		});
		excelTitleMap.forEach((k, v) -> {
			Map<Integer, List<ExcelTemplateTitleModel>> titleMap = v.stream().collect(Collectors.groupingBy(ExcelTemplateTitleModel::getRowIndex));
			templateTitleCacheMap.put(k, titleMap);
		});

	}

	/**
	 * 加载导入模板
	 *
	 * @param templates                       模板文件list
	 * @param excelImportTemplateFileCacheMap 模板文件缓存map
	 * @param excelImportTemplateCacheMap     模板缓存map
	 * @param excluder                        模板缓存排除器
	 */
	public static void loadBootImportTemplate(List<Map<String, Object>> templates, Map<String, byte[]> excelImportTemplateFileCacheMap, Map<String, Map<String, List<ExcelImportTemplateCacheModel>>> excelImportTemplateCacheMap, AbstractImportTemplateExcluder excluder) {
		if (templates.isEmpty()) {
			return;
		}
		ExecutorService executorService = new ThreadPoolExecutor(templates.size(), templates.size(), ExcelConstant.INT_1000, TimeUnit.MILLISECONDS, new SynchronousQueue<>(), new ThreadFactory() {
			int i = ExcelConstant.ZERO_SHORT;

			@Override
			public Thread newThread(Runnable r) {
				i++;
				return new Thread(r, LOAD_IMPORT_TEMPLATE_POOL_PREFIX + i);
			}
		});
		CountDownLatch countDownLatch = new CountDownLatch(templates.size());
		List<TaskExcuteResult> futureList = new ArrayList<>(templates.size());
		for (Map<String, Object> template : templates) {
			String fileName = (String)template.get("name");
			InputStream inputStream = (InputStream)template.get("input");
			InputStream cacheInput = (InputStream)template.get("cacheInput");
			Future<Boolean> future = executorService.submit(new LoadImportTemplateTask(inputStream, cacheInput, fileName, excelImportTemplateFileCacheMap, countDownLatch, excelImportTemplateCacheMap, excluder));
			TaskExcuteResult excuteResult = TaskExcuteResult.builder().result(future).templateName(fileName).build();
			futureList.add(excuteResult);
		}
		try {
			countDownLatch.await();
			for (TaskExcuteResult excuteResult : futureList) {
				log.info("Load excel import template {} by {} ", excuteResult.result.get() ? "successfully" : "failed", excuteResult.templateName);
			}
		} catch (InterruptedException e) {
			log.error("Load excel import template failed,cause:{}", Throwables.getStackTraceAsString(e));
		} catch (ExecutionException e) {
			log.error("Load excel import template failed,cause:{}", Throwables.getStackTraceAsString(e));
		} finally {
			if (Objects.nonNull(executorService)) {
				executorService.shutdown();
			}
		}
	}


	/**
	 * @Description: 加载导出模板线程任务
	 * @Author: Vachel Wang
	 * @Date: 2019/9/26 上午9:55
	 */
	private static class LoadExportTemplateTask implements Callable<Boolean> {
		private InputStream inputStream;
		private InputStream cacheInput;
		private String excelFileName;
		private List<ExcelExpressionModel> expressionModelList;
		private List<ExcelTemplateTitleModel> titleModelList;
		private Pattern pattern = Pattern.compile(ExcelConstant.EXPRESSION_PATTERN);
		private Map<String, ExcelTemplateCacheModel> excelTemplateFileCacheMap;

		public LoadExportTemplateTask(InputStream inputStream, InputStream cacheInput, String excelFileName, List<ExcelExpressionModel> expressionModelList, List<ExcelTemplateTitleModel> titleModelList, Map<String, ExcelTemplateCacheModel> excelTemplateFileCacheMap) {
			this.inputStream = inputStream;
			this.cacheInput = cacheInput;
			this.excelFileName = excelFileName;
			this.expressionModelList = expressionModelList;
			this.titleModelList = titleModelList;
			this.excelTemplateFileCacheMap = excelTemplateFileCacheMap;
		}

		@Override
		public Boolean call() {
			try (InputStream inputStream = this.inputStream; InputStream cacheInputSteam = this.cacheInput; Workbook workbook = this.excelFileName.endsWith(ExcelConstant.XLSX_STR) ? new XSSFWorkbook(inputStream) : new HSSFWorkbook(inputStream)) {
				for (Sheet sheet : workbook) {
					for (Row row : sheet) {
						for (Cell cell : row) {
							cell.setCellType(CellType.STRING);
							String value = cell.getStringCellValue();
							if (StringUtil.notEmpty(value)) {
								Matcher matcher = pattern.matcher(value);
								boolean hasExpression = false;
								while (matcher.find()) {
									hasExpression = true;
									String expression = matcher.group(ExcelConstant.ZERO_SHORT);
									String expressionContent = expression.substring(ExcelConstant.TOW_INT, expression.length() - ExcelConstant.ONE_INT);
									int dotIndex = expressionContent.indexOf(ExcelConstant.DOT_CHAR);
									String nameSpace = expressionContent.substring(ExcelConstant.ZERO_SHORT, dotIndex);
									String[] fieldName = expressionContent.substring(dotIndex + ExcelConstant.ONE_INT).split(ExcelConstant.DOT_REGEX);

									ExcelExpressionModel expressionModel = ExcelExpressionModel.builder().colIndex(cell.getColumnIndex()).rowIndex(row.getRowNum()).sheetName(sheet.getSheetName()).expressionContent(expressionContent).expression(expression).fieldName(fieldName).nameSpace(nameSpace).build();
									expressionModelList.add(expressionModel);
								}
								// 标题
								if (!hasExpression) {
									int prevCellIndex = cell.getColumnIndex() - ExcelConstant.ONE_INT;
									if (prevCellIndex < ExcelConstant.ZERO_SHORT) {
										Cell nextCell = row.getCell(cell.getColumnIndex() + ExcelConstant.ONE_INT);
										String nextCellValue = null;
										if (Objects.nonNull(nextCell)) {
											nextCell.setCellType(CellType.STRING);
											nextCellValue = nextCell.getStringCellValue();
										}

										boolean nextIsExpression = Optional.ofNullable(nextCellValue).orElse(ExcelConstant.NULL_STR).matches(ExcelConstant.EXPRESSION_TEXT_PATTERN);
										if (!nextIsExpression && StringUtil.notEmpty(nextCellValue)) {
											ExcelTemplateTitleModel titleModel = ExcelTemplateTitleModel.builder().title(value).colIndex(cell.getColumnIndex()).rowIndex(row.getRowNum()).sheetName(sheet.getSheetName()).build();
											titleModelList.add(titleModel);
										}
									} else {
										Cell prevCell = row.getCell(prevCellIndex), nextCell = row.getCell(cell.getColumnIndex() + ExcelConstant.ONE_INT);
										String prevCellValue = null, nextCellValue = null;
										if (Objects.nonNull(nextCell)) {
											nextCell.setCellType(CellType.STRING);
											nextCellValue = nextCell.getStringCellValue();
										}
										if (Objects.nonNull(prevCell)) {
											prevCell.setCellType(CellType.STRING);
											prevCellValue = prevCell.getStringCellValue();
										}
										boolean prevIsExpression = Optional.ofNullable(prevCellValue).orElse(ExcelConstant.NULL_STR).matches(ExcelConstant.EXPRESSION_TEXT_PATTERN), nextIsExpression = Optional.ofNullable(nextCellValue).orElse(ExcelConstant.NULL_STR).matches(ExcelConstant.EXPRESSION_TEXT_PATTERN);

										if ((!prevIsExpression && StringUtil.notEmpty(prevCellValue)) || (!nextIsExpression && StringUtil.notEmpty(nextCellValue))) {
											ExcelTemplateTitleModel titleModel = ExcelTemplateTitleModel.builder().title(value).colIndex(cell.getColumnIndex()).rowIndex(row.getRowNum()).sheetName(sheet.getSheetName()).build();
											titleModelList.add(titleModel);
										}
									}
								}
							}
						}
					}
				}
				Workbook cacheWorkbook = excelFileName.endsWith(ExcelConstant.XLSX_STR) ? new XSSFWorkbook(cacheInputSteam) : new HSSFWorkbook(cacheInputSteam);
				ThreadLocal<WorkbookCachePool.WorkbookCacheModel> workbookCacheModelThreadLocal = WorkbookCachePool.addBasicStyle(cacheWorkbook);
				ExcelTemplateCacheModel cacheModel = new ExcelTemplateCacheModel(workbookCacheModelThreadLocal);
				excelTemplateFileCacheMap.put(excelFileName, cacheModel);
			} catch (Exception e) {
				log.error("Load excel template failed,cause:{}", Throwables.getStackTraceAsString(e));
				return Boolean.FALSE;
			}
			return Boolean.TRUE;
		}
	}

	/**
	 * @Description: 加载导入模板任务
	 * @Author: Vachel Wang
	 * @Date: 2020/1/11 下午4:20
	 * @Email:
	 */
	private static class LoadImportTemplateTask implements Callable<Boolean> {
		private String fileName;
		private InputStream inputStream;
		private InputStream cacheInput;
		private Map<String, byte[]> excelImportTemplateFileCacheMap;
		private CountDownLatch countDownLatch;
		private Map<String, Map<String, List<ExcelImportTemplateCacheModel>>> excelImportTemplateCacheMap;
		private AbstractImportTemplateExcluder excluder;

		public LoadImportTemplateTask(InputStream inputStream, InputStream cacheInput, String fileName, Map<String, byte[]> excelImportTemplateFileCacheMap, CountDownLatch countDownLatch, Map<String, Map<String, List<ExcelImportTemplateCacheModel>>> excelImportTemplateCacheMap, AbstractImportTemplateExcluder excluder) {
			this.inputStream = inputStream;
			this.cacheInput = cacheInput;
			this.fileName = fileName;
			this.excelImportTemplateFileCacheMap = excelImportTemplateFileCacheMap;
			this.countDownLatch = countDownLatch;
			this.excelImportTemplateCacheMap = excelImportTemplateCacheMap;
			this.excluder = excluder;
		}

		@Override
		public Boolean call() {

			try (InputStream templateCacheInput = this.cacheInput; InputStream templateInput = this.inputStream; Workbook workbook = fileName.endsWith(ExcelConstant.XLSX_STR) ? new XSSFWorkbook(templateInput) : new HSSFWorkbook(templateInput)) {
				Map<String, List<ExcelImportTemplateCacheModel>> sheetCacheMap = new HashMap<>();
				for (Sheet sheet : workbook) {
					String sheetName = sheet.getSheetName();
					if (Objects.nonNull(excluder) && excluder.isExclude(excluder.new SheetExclude(sheetName, fileName))) {
						continue;
					}
					List<ExcelImportTemplateCacheModel> cacheModelList = new ArrayList<>();
					for (Row row : sheet) {
						if (Objects.nonNull(excluder) && excluder.isExclude(excluder.new RowExclude(row.getRowNum(), sheetName, fileName))) {
							continue;
						}
						for (Cell cell : row) {
							if (Objects.nonNull(excluder) && excluder.isExclude(excluder.new ColumnExclude(cell.getColumnIndex(), cell.getRowIndex(), sheetName, fileName))) {
								continue;
							}
							cell.setCellType(CellType.STRING);
							String text = cell.getStringCellValue();
							if (StringUtil.isEmpty(text)) {
								continue;
							}
							ExcelImportTemplateCacheModel cacheModel = new ExcelImportTemplateCacheModel();
							cacheModel.setColIndex(cell.getColumnIndex());
							cacheModel.setRowIndex(cell.getRowIndex());
							cacheModel.setText(cell.getStringCellValue());
							cacheModelList.add(cacheModel);
						}
					}
					sheetCacheMap.put(sheet.getSheetName(), cacheModelList);
				}
				excelImportTemplateCacheMap.put(fileName, sheetCacheMap);
				excelImportTemplateFileCacheMap.put(fileName, IOUtils.toByteArray(templateCacheInput));
			} catch (Exception e) {
				log.error("Load excel template failed,cause:{}", Throwables.getStackTraceAsString(e));
				return Boolean.FALSE;
			} finally {
				countDownLatch.countDown();
			}
			return Boolean.TRUE;
		}
	}

	/**
	 * 任务执行结果
	 */
	@Data
	@Builder
	static class TaskExcuteResult {
		private Future<Boolean> result;
		private String templateName;
	}
}
