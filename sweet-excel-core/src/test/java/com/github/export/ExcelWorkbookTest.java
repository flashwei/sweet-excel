package com.github.export;

import lombok.Data;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Future;
import java.util.concurrent.SynchronousQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

/**
 * @description: excel 异步导出测试
 * @author: Vachel Wang
 * @create: 2019-12-03 16:13
 **/
public class ExcelWorkbookTest {

	@Test
	public void testAsyncExport() throws Exception{
		XSSFWorkbook workbook = new XSSFWorkbook();
		ExecutorService executorService = new ThreadPoolExecutor(3, 10, 3, TimeUnit.SECONDS, new SynchronousQueue<>());
		List<Future<Boolean>> futureList = new ArrayList<>();

		int rowIndex = 0 ;
		for(int i=1;i<=5;i++) {
			List<DataDto> dataList = new ArrayList<>();
			for (int j = 0; j <= 200; j++) {
				DataDto dto = new DataDto();
				dto.setRowIndex(j);
				dto.setValue("data-" + i + "-" + j);
				dataList.add(dto);
				rowIndex++;
			}
			Bank bank = new Bank();
			AsyncExportTask exportTask = new AsyncExportTask(dataList,workbook,bank);
			Future<Boolean> future = executorService.submit(exportTask);
			futureList.add(future);
		}
		for (Future<Boolean> future : futureList) {
			while (!future.isDone()) {

			}

		}

	}


	class AsyncExportTask implements Callable<Boolean> {
		private List<DataDto> dataList;
		private final ThreadLocal<Workbook> workbookThreadLocal  ;
		private Bank bank ;
		private Lock lock = new ReentrantLock();

		public AsyncExportTask(List<DataDto> dataList, final Workbook workbook,Bank bank){
			workbookThreadLocal = ThreadLocal.withInitial(()->{return workbook;});
			this.dataList = dataList;
			this.bank = bank ;
		}
		@Override
		public Boolean call() throws Exception {
			try {
				Workbook workbook = workbookThreadLocal.get();
				int num = 1000 + (int)(Math.random() * (9999-1000+1));
				String sheetName = "sheet"+num+""+System.currentTimeMillis();
				lock.lock();
				bank.saveBalance(100);
				System.out.println(sheetName);
				System.out.println(bank.getBalance());
				lock.unlock();
				/*Sheet sheet = workbook.createSheet(sheetName);
				for (DataDto dto : dataList) {

					try {
						Row row = sheet.createRow(dto.getRowIndex());
						Cell cell = row.createCell(0);
						cell.setCellValue(dto.getValue());
					} catch (Exception e) {
						e.printStackTrace();
					}

				}
				Sheet test = workbook.getSheet(sheetName);
				Row row = test.getRow(1);
				Cell cell = row.getCell(0);
				System.out.println(Thread.currentThread().getName()+","+cell.getStringCellValue()+","+System.identityHashCode(workbook));*/


//				workbook.write(outputStream);
			} catch (Exception e) {
				e.printStackTrace();
			}finally {

			}
			return true;
		}

	}
	@Data
	class DataDto{
		private Integer rowIndex ;
		private String value ;
	}


}

@Data
class WorkBookModel implements Serializable{
	private byte[] bytes ;
}

class Bank{
	private volatile int balance = 0 ;

	public void saveBalance(int num) {
		Lock lock = new ReentrantLock();
		lock.lock();
		balance+=num;
		lock.unlock();
	}

	public int getBalance() {
		return balance;
	}
}