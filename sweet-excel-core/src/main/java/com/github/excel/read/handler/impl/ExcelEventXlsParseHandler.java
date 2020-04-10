package com.github.excel.read.handler.impl;

import com.github.excel.read.ExcelEventReader;
import com.github.excel.read.ExcelEventRowReader;
import com.google.common.base.Throwables;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.ExcelEventReadExecutor;
import com.github.excel.read.handler.AbstractExecuteHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:11 下午
 * @Description: 事件模式下解析xls处理器
 */
@Slf4j
@SuppressWarnings("unchecked")
public class ExcelEventXlsParseHandler<T> implements HSSFListener, ExcelEventReader<T> {

	private int minColumns = -1;

	private POIFSFileSystem fs;

	private int lastRowNumber;

	private int lastColumnNumber;

	/** Should we output the formula, or the value it has? */
	private boolean outputFormulaValues = true;

	/** For parsing Formulas */
	private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;

	/**
	 * excel2003工作薄
	 */
	private HSSFWorkbook stubWorkbook;

	/**
	 * Records we pick up as we process
	 */
	private SSTRecord sstRecord;

	private FormatTrackingHSSFListener formatListener;

	// 表索引
	private int sheetIndex = -1;

	private BoundSheetRecord[] orderedBSRs;

	@SuppressWarnings("unchecked")
	private ArrayList boundSheetRecords = new ArrayList();

	/**
	 * For handling formulas with string results
	 */
	private int nextRow;

	private int nextColumn;

	private boolean outputNextStringRecord;

	// 当前行
	private int curRow = 0;

	/**
	 * 存储行记录的容器
	 */
	private List<String> rowlist = new ArrayList<>();

	@SuppressWarnings("unused")
	private String sheetName;

	private ExcelEventRowReader<T> rowReader;
	private ExcelEventReadExecutor<T> excelEventReadExecutor;

	@Override
	public void setRowReader(ExcelEventRowReader<T> rowReader) {
		this.rowReader = rowReader;
	}

	@Override
	public void setExecuteHandler(AbstractExecuteHandler<T> executeHandler){
		this.excelEventReadExecutor = new ExcelEventReadExecutor<>(executeHandler);
	}

	/**
	 * 遍历excel下所有的sheet
	 *
	 * @throws IOException
	 */
	@Override
	public void process(InputStream inputStream) throws ExcelReadException {
		try {
			this.fs = new POIFSFileSystem(inputStream);
			MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
			formatListener = new FormatTrackingHSSFListener(listener);
			HSSFEventFactory factory = new HSSFEventFactory();
			HSSFRequest request = new HSSFRequest();
			if (outputFormulaValues) {
				request.addListenerForAllRecords(formatListener);
			} else {
				workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
				request.addListenerForAllRecords(workbookBuildingListener);
			}
			factory.processWorkbookEvents(request, fs);
			excelEventReadExecutor.flush();
		} catch (ExcelReadException e) {
			log.error("ExcelException:{}", Throwables.getStackTraceAsString(e));
			throw e;
		} catch (IOException e) {
			log.error("IOException:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException("io.exception");
		} catch (Exception e) {
			log.error("Exception:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException("exception");
		}
	}

	/**
	 * HSSFListener 监听方法，处理 Record
	 */
	@SuppressWarnings("unchecked")
	@Override
	public void processRecord(Record record) {
		int thisRow = -1;
		int thisColumn = -1;
		String thisStr = null;
		String value = null;
		switch (record.getSid()) {
			case BoundSheetRecord.sid:
				boundSheetRecords.add(record);
				break;
			case BOFRecord.sid:
				BOFRecord br = (BOFRecord) record;
				if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
					// 如果有需要，则建立子工作薄
					if (workbookBuildingListener != null && stubWorkbook == null) {
						stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
					}

					sheetIndex++;
					if (orderedBSRs == null) {
						orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
					}
					sheetName = orderedBSRs[sheetIndex].getSheetname();
				}
				break;

			case SSTRecord.sid:
				sstRecord = (SSTRecord) record;
				break;

			case BlankRecord.sid:
				BlankRecord brec = (BlankRecord) record;
				thisRow = brec.getRow();
				thisColumn = brec.getColumn();
				thisStr = "";
				rowlist.add(thisColumn, thisStr);
				break;
			case BoolErrRecord.sid: // 单元格为布尔类型
				BoolErrRecord berec = (BoolErrRecord) record;
				thisRow = berec.getRow();
				thisColumn = berec.getColumn();
				thisStr = berec.getBooleanValue() + "";
				rowlist.add(thisColumn, thisStr);
				break;

			case FormulaRecord.sid: // 单元格为公式类型
				FormulaRecord frec = (FormulaRecord) record;
				thisRow = frec.getRow();
				thisColumn = frec.getColumn();
				if (outputFormulaValues) {
					if (Double.isNaN(frec.getValue())) {
						// Formula result is a string
						// This is stored in the next record
						outputNextStringRecord = true;
						nextRow = frec.getRow();
						nextColumn = frec.getColumn();
					} else {
						thisStr = formatListener.formatNumberDateCell(frec);
					}
				} else {
					thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
				}
				rowlist.add(thisColumn, thisStr);
				break;
			case StringRecord.sid:// 单元格中公式的字符串
				if (outputNextStringRecord) {
					// String for formula
					StringRecord srec = (StringRecord) record;
					thisStr = srec.getString();
					thisRow = nextRow;
					thisColumn = nextColumn;
					outputNextStringRecord = false;
				}
				break;
			case LabelRecord.sid:
				LabelRecord lrec = (LabelRecord) record;
				curRow = thisRow = lrec.getRow();
				thisColumn = lrec.getColumn();
				value = lrec.getValue().trim();
				value = value.equals("") ? " " : value;
				this.rowlist.add(thisColumn, value);
				break;
			case LabelSSTRecord.sid: // 单元格为字符串类型
				LabelSSTRecord lsrec = (LabelSSTRecord) record;
				curRow = thisRow = lsrec.getRow();
				thisColumn = lsrec.getColumn();
				if (sstRecord == null) {
					rowlist.add(thisColumn, " ");
				} else {
					value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
					value = value.equals("") ? " " : value;
					rowlist.add(thisColumn, value);
				}
				break;
			case NumberRecord.sid: // 单元格为数字类型
				NumberRecord numrec = (NumberRecord) record;
				curRow = thisRow = numrec.getRow();
				thisColumn = numrec.getColumn();
				value = formatListener.formatNumberDateCell(numrec).trim();
				value = value.equals("") ? " " : value;
				// 向容器加入列值
				rowlist.add(thisColumn, value);
				break;
			default:
				break;
		}

		// 遇到新行的操作
		if (thisRow != -1 && thisRow != lastRowNumber) {
			lastColumnNumber = -1;
		}

		// 空值的操作
		if (record instanceof MissingCellDummyRecord) {
			MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
			curRow = thisRow = mc.getRow();
			thisColumn = mc.getColumn();
			rowlist.add(thisColumn, " ");
		}

		// 更新行和列的值
		if (thisRow > -1) {
			lastRowNumber = thisRow;
		}
		if (thisColumn > -1) {
			lastColumnNumber = thisColumn;
		}

		// 行结束时的操作
		if (record instanceof LastCellOfRowDummyRecord) {
			if (minColumns > 0) {
				// 列值重新置空
				if (lastColumnNumber == -1) {
					lastColumnNumber = 0;
				}
			}
			lastColumnNumber = -1;

			// 每行结束时，调用getRows() 方法
			T object = rowReader.getRows(sheetIndex, curRow, rowlist);
			//将对象交给执行处理器处理
			if (Objects.nonNull(object)) {
				excelEventReadExecutor.submit(object);
			}
			// 清空容器
			rowlist.clear();
		}
	}


	public ExcelEventXlsParseHandler addRowReader(ExcelEventRowReader rowReader) {
		setRowReader(rowReader);
		return this;
	}


	public ExcelEventXlsParseHandler addExecuteHandler(AbstractExecuteHandler executeHandler){
		setExecuteHandler(executeHandler);
		return this;
	}
}
