package com.github.export;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Future;
import java.util.concurrent.SynchronousQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:26 下午
 * @Description: 多线程导出
 */
public class ExcelMutiliThreadTest {

	private ThreadLocal<Workbook> workbookThreadLocal = ThreadLocal.withInitial(()->{return new HSSFWorkbook();});
	@Test
	public void testMutiliThread() {
		ExecutorService executorService = new ThreadPoolExecutor(2, 3, 1000, TimeUnit.SECONDS, new SynchronousQueue<>());
		List<Future> futureList = new ArrayList<>();
		for (int i = 1; i <= 3; i++) {
			Future future = executorService.submit(new Task());
			futureList.add(future);
		}

		for (Future future : futureList) {
			while(!future.isDone()){}
		}
	}

	class Task implements Runnable{
		@Override
		public void run() {
			try(Workbook workbook = workbookThreadLocal.get()) {
				Sheet sheet = workbook.createSheet("sheet");
				Row row = sheet.createRow(0);
				Cell cell = row.createCell(0);
				cell.setCellValue("abc");
				System.out.println(cell.getStringCellValue());
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
