package com.github.excel.read.handler.impl;

import com.github.excel.read.ExcelEventReader;
import com.github.excel.read.ExcelEventRowReader;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.ExcelEventReadExecutor;
import com.github.excel.read.handler.AbstractExecuteHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:11 下午
 * @Description: 事件模式下解析xlsx处理器
 */
@Slf4j
public class ExcelEventXlsxParseHandler<T> extends DefaultHandler implements ExcelEventReader<T> {

	/**
	 * 共享字符串表
	 */
	private SharedStringsTable sst;

	/**
	 * 上一次的内容
	 */
	private StringBuilder lastContents;

	/**
	 * 工作表索引
	 */
	private int sheetIndex = -1;

	/**
	 * 行数据，存储字符串格式
	 * 需要在handler中自行转换
	 */
	private List<String> currentRowValues;

	/**
	 * T元素标识
	 */
	private boolean isTElement;

	/**
	 * 有效数据矩形区域,A1:Y2
	 */
	private String dimension;

	/**
	 * 根据dimension得出每行的数据长度
	 */
	private int longest;

	/**
	 * 表格的读取的当前行数
	 */
	private Integer curRowId;

	/**
	 * 上个有内容的单元格id，判断空单元格
	 */
	private String lastColId;

	private Map<String, Boolean> existCols;

	private ExcelEventRowReader<T> rowReader;

	private ExcelEventReadExecutor<T> excelEventReadExecutor;

	@Override
	public void setRowReader(ExcelEventRowReader<T> rowReader){
		this.rowReader = rowReader;
	}

	@Override
	public void setExecuteHandler(AbstractExecuteHandler<T> executeHandler){
		this.excelEventReadExecutor = new ExcelEventReadExecutor<>(executeHandler);
	}

	/**
	 * 遍历工作簿中所有的电子表格
	 * @param inputStream
	 * @throws Exception
	 */
	@Override
	public void process(InputStream inputStream) throws ExcelReadException {
		try {
			OPCPackage pkg = OPCPackage.open(inputStream);
			XSSFReader r = new XSSFReader(pkg);
			SharedStringsTable sst = r.getSharedStringsTable();
			XMLReader parser = fetchSheetParser(sst);
			Iterator<InputStream> sheets = r.getSheetsData();
			while (sheets.hasNext()) {
				sheetIndex++;
				InputStream sheet = sheets.next();
				InputSource sheetSource = new InputSource(sheet);
				parser.parse(sheetSource);
				excelEventReadExecutor.flush();
				sheet.close();
			}
		} catch (ExcelReadException e){
			log.error("ExcelException:{}", Throwables.getStackTraceAsString(e));
			throw e;
		} catch (IOException e) {
			log.error("IOException:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException("io.exception");
		} catch (Exception e) {
			log.error("IOException:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelReadException("io.exception");
		}
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst)
			throws SAXException, ParserConfigurationException {
		SAXParserFactory m_parserFactory = SAXParserFactory.newInstance();
		m_parserFactory.setNamespaceAware(true);
		XMLReader parser = m_parserFactory.newSAXParser().getXMLReader();
		this.sst = sst;
		parser.setContentHandler(this);
		return parser;
	}

	@Override
	public void startElement(String uri, String localName, String name,
							 Attributes attributes) throws SAXException {

		if (name.equals("dimension")){
			/**
			 * <dimension ref="A1:E3" />
			 * 有效取值范围：第一行的A列开始，到第三行的E列
			 */
			dimension = attributes.getValue("ref");
			// 最大列数
			longest = covertRowIdtoInt(dimension.substring(dimension.indexOf(":")+1));
		}

		// 一行开始
		if (name.equals("row")) {
			// <row r="1" spans="1:5" x14ac:dyDescent="0.15">...</row>
			currentRowValues = Lists.newArrayList();
			existCols = Maps.newHashMap();
			curRowId = Integer.valueOf(attributes.getValue("r"));
		}

		// 一行里的单元格开始
		if ("c".equals(name)) {
			// <c r="A1" s="1" t="s"><v>hello</v></c>
			String curColId = attributes.getValue("r");
			existCols.put(curColId, false);

			//判断同一行的空单元，取空字符串放入currentRowValues
			if (lastColId !=null) {
				int gap = covertRowIdtoInt(curColId)-covertRowIdtoInt(lastColId);
				for(int i=0;i<gap-1;i++) {
					currentRowValues.add("");
				}
			} else {
				//第一个单元格可能不是在第一列
				if (!"A1".equals(curColId)) {
					for(int i = 0; i < covertRowIdtoInt(curColId)-1; i++) {
						currentRowValues.add("");
					}
				}
			}

			lastColId = curColId;

			//判断单元格的值是SST 的索引，不能直接characters方法取值
			if (attributes.getValue("t")!=null && attributes.getValue("t").equals("s")) {
				isTElement = true;
			}else{
				isTElement = false;
			}
		}

		// 置空
		lastContents = new StringBuilder("");
	}

	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {

		//如果标签名称为 row ，说明已到行尾
		if (name.equals("row")) {

			//判断最后一个单元格是否在最后，补齐列数
			if (Objects.nonNull(lastColId)) {
				if (covertRowIdtoInt(lastColId) < longest) {
					for (int i = 0; i < longest - covertRowIdtoInt(lastColId); i++) {
						currentRowValues.add("");
					}
				}
			}

			// 将对象交给执行处理器处理
			Object o = rowReader.getRows(sheetIndex, curRowId, currentRowValues);
			T object = (T)o;
			excelEventReadExecutor.submit(object);
			currentRowValues.clear();
			lastColId = null;
		}

		//单元格内容标签结束，characters方法会被调用处理内容
		if (name.equals("v")) {
			//单元格的值是SST 的索引
			if (isTElement){
				try {
					int idx = Integer.parseInt(lastContents.toString());
					lastContents = new StringBuilder(new XSSFRichTextString(sst.getEntryAt(idx)).toString());
					currentRowValues.add(lastContents.toString());
					existCols.put(lastColId, true);
				} catch (Exception e) {

				}
			}else {
				currentRowValues.add(lastContents.toString());
				existCols.put(lastColId, true);
			}
		}

		if(name.equals("t")){
			currentRowValues.add(lastContents.toString());
			existCols.put(lastColId, true);
		}

		if ("c".equals(name)) {
			// 读取了列但是未赋值
			if (existCols.containsKey(lastColId) && !existCols.get(lastColId)) {
				currentRowValues.add(lastContents.toString());
				existCols.put(lastColId, true);
			}
		}

	}

	@Override
	public void characters(char[] ch, int start, int length)
			throws SAXException {
		//得到单元格内容的值
		lastContents.append(ch,start,length);
	}

	/**
	 * 列号转数字   AB7-->28 第28列
	 * @param rowId
	 * @return
	 */
	public static int covertRowIdtoInt(String rowId){
		int firstDigit = -1;
		for (int c = 0; c < rowId.length(); ++c) {
			if (Character.isDigit(rowId.charAt(c))) {
				firstDigit = c;
				break;
			}
		}
		//AB7-->AB
		//AB是列号, 7是行号
		String newRowId = rowId.substring(0,firstDigit);
		int num = 0;
		int result = 0;
		int length = newRowId.length();
		for(int i = 0; i < length; i++) {
			//先取最低位，B
			char ch = newRowId.charAt(length - i - 1);
			//B表示的十进制2，ascii码相减，以A的ascii码为基准，A表示1，B表示2
			num = (int)(ch - 'A' + 1) ;
			//列号转换相当于26进制数转10进制
			num *= Math.pow(26, i);
			result += num;
		}
		return result;

	}

	public ExcelEventXlsxParseHandler addRowReader(ExcelEventRowReader rowReader) {
		setRowReader(rowReader);
		return this;
	}


	public ExcelEventXlsxParseHandler addExecuteHandler(AbstractExecuteHandler executeHandler){
		setExecuteHandler(executeHandler);
		return this;
	}
}
