package com.github.export;

import com.google.common.base.Stopwatch;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.*;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:25 下午
 * @Description: excel 异步导出测试
 */
public class ExcelAsyncExoprtTest {

	@Test
	public void testAsyncExport() throws Exception{
		SXSSFWorkbook workbook = new SXSSFWorkbook(-1);
		SXSSFSheet sheet = workbook.createSheet("test");
		ExecutorService executorService = new ThreadPoolExecutor(3, 10, 3, TimeUnit.SECONDS, new SynchronousQueue<>());
		List<Future<Boolean>> futureList = new ArrayList<>();
		int rowIndex = 0 ;
		for(int i=1;i<=10;i++) {
			List<DataDto> dataList = new ArrayList<>();
			for (int j = 0; j <= 200; j++) {
				DataDto dto = new DataDto();
				dto.setRowIndex(rowIndex);
				dto.setValue("data-" + i + "-" + j);
				dataList.add(dto);
				rowIndex++;
			}
			AsyncExportTask exportTask = new AsyncExportTask(sheet,dataList,workbook);
			Future<Boolean> future = executorService.submit(exportTask);
			futureList.add(future);
		}
		for (Future<Boolean> future : futureList) {
			while (!future.isDone()) {

			}
		}
		OutputStream outputStream = new FileOutputStream(new File("/Users/vachel/Documents/test-111.xlsx"));
		workbook.write(outputStream);
	}


	class AsyncExportTask implements Callable<Boolean> {
		private Sheet sheet ;
		private List<DataDto> dataList;
		private Workbook workbook ;
		public AsyncExportTask(Sheet sheet , List<DataDto> dataList,Workbook workbook){
			this.sheet = sheet;
			this.dataList = dataList;
			this.workbook = workbook;
		}

		@Override
		public Boolean call() throws Exception {
			for (DataDto dto : dataList) {

				try {
					System.out.println(sheet);
					System.out.println(dto);
					Row row = this.sheet.createRow(dto.getRowIndex());
					Cell cell = row.createCell(0);
					cell.setCellValue(dto.getValue());
					System.out.println(dto.getValue());
				} catch (Exception e) {
					e.printStackTrace();
				}

			}
			Sheet testSheet = workbook.getSheet("test22");
			if (Objects.isNull(testSheet)) {
				workbook.createSheet("test22");
			}
			return true;
		}
	}
	@Data
	class DataDto{
		private Integer rowIndex ;
		private String value ;
	}
	@Test
	public void testRead() throws Exception{
		FileInputStream fileInputStream = new FileInputStream(new File("/Users/vachel/Documents/test-error.xlsx"));
		byte[] bytes = new byte[fileInputStream.available()];
		fileInputStream.read(bytes);
		int count = 5 ;
		long totalMillise = 0 ;
		for(int k=1;k<=count;k++) {
			Stopwatch stopwatch = Stopwatch.createStarted();
			for (int i = 1; i <= 1000; i++) {
				ByteArrayInputStream inputStream = new ByteArrayInputStream(bytes);
				XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
				inputStream.close();
			}
			long elapsed = stopwatch.elapsed(TimeUnit.MILLISECONDS);
			totalMillise += elapsed ;
			System.out.println("ByteArrayInputStream:" + elapsed);
		}
		System.out.println(totalMillise/count);
/*
		fileInputStream = new FileInputStream(new File("/Users/vachel/Documents/test-error.xlsx"));
		stopwatch = Stopwatch.createStarted();
		workbook = new XSSFWorkbook(fileInputStream);
		elapsed = stopwatch.elapsed(TimeUnit.MILLISECONDS);
		System.out.println("fileInputStream:"+elapsed);

		XSSFSheet sheetAt = workbook.getSheetAt(0);
		System.out.println(sheetAt.getRow(0).getCell(0).getStringCellValue());
		System.out.println(sheetAt.getRow(0).getCell(1).getStringCellValue());
		System.out.println("=======================");
		inputStream.reset();
		workbook = new XSSFWorkbook(inputStream);
		sheetAt = workbook.getSheetAt(0);
		System.out.println(sheetAt.getRow(0).getCell(0).getStringCellValue());
		System.out.println(sheetAt.getRow(0).getCell(1).getStringCellValue());*/

		/*for (int i = 0; i < sheetAt.getPhysicalNumberOfRows(); i++) {
			Row row = sheetAt.getRow(i);
			for (int j = 0; j <= row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
				System.out.println(cell);
			}
		}*/
	}
	@Test
	public void testRead1() throws Exception{
		Stopwatch stopwatch = Stopwatch.createStarted();
		int count = 5 ;
		long totalMillise = 0 ;
		for(int k=1;k<=count;k++) {
			for (int i = 1; i <= 1000; i++) {
				FileInputStream fileInputStream = new FileInputStream(new File("/Users/vachel/Documents/test-error.xlsx"));
				XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
				fileInputStream.close();
			}
			long elapsed = stopwatch.elapsed(TimeUnit.MILLISECONDS);
			totalMillise += elapsed ;
			System.out.println("fileInputStream:" + elapsed);
		}
		System.out.println(totalMillise/count);
		/*XSSFSheet sheetAt = workbook.getSheetAt(0);
		System.out.println(sheetAt.getRow(0).getCell(0).getStringCellValue());
		System.out.println(sheetAt.getRow(0).getCell(1).getStringCellValue());
		System.out.println("=======================");*/
		/*inputStream.reset();
		workbook = new XSSFWorkbook(inputStream);
		sheetAt = workbook.getSheetAt(0);
		System.out.println(sheetAt.getRow(0).getCell(0).getStringCellValue());
		System.out.println(sheetAt.getRow(0).getCell(1).getStringCellValue());*/

		/*for (int i = 0; i < sheetAt.getPhysicalNumberOfRows(); i++) {
			Row row = sheetAt.getRow(i);
			for (int j = 0; j <= row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
				System.out.println(cell);
			}
		}*/
	}

	@Test
	public void test() throws Exception{

		WorkBookModel model = new WorkBookModel();
		String str = "11";
		model.setBytes(new byte[2048]);
		model.getBytes()[1]=1;
//		model.setWorkbook(new XSSFWorkbook());
//		model.setName("jack");
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		ObjectOutputStream oos = new ObjectOutputStream(bos);
		oos.writeObject(model);

//		ByteArrayInputStream inputStream = new ByteArrayInputStream(bos.write)
	}
	@Test
	public void test1() throws Exception{

		ObjectInputStream inputStream = new ObjectInputStream(new FileInputStream("/Users/vachel/Documents/workbook.ser"));
		Object o = inputStream.readObject();
		WorkBookModel model = (WorkBookModel)o;
		System.out.println(1);
	}


}
