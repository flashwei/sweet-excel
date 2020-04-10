package com.github.export;

import com.github.model2.CompanyDto2;
import com.google.common.base.Stopwatch;
import com.google.common.collect.Lists;
import com.github.BaseExcelTest;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.util.ZipCompressUtil;
import com.github.excel.write.ExcelLargeListBatchWriter;
import com.github.excel.write.ExcelLargeListWriter;
import com.github.excel.write.ExcelWriterFactory;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;


/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:26 下午
 * @Description: 大文本导出
 */
@Slf4j
public class ExcelLargeListWriterTest extends BaseExcelTest {

	@Test
	public void testExport() throws Exception {
		ExcelLargeListWriter listWriter = ExcelWriterFactory.createLargeListMultiThreadWriter("数据导出", 2,100, CompanyDto2.class);
		listWriter.addStyle(ExcelCustomStyle.class);
		List<CompanyDto2> companyDtoList = Lists.newArrayList();
		for (int i = 1; i <= 100; i++) {
			CompanyDto2 company = new CompanyDto2();
			company.setAddress("address" + i);
			company.setName(i + "-短笛阿德南科技");
			company.setCreateTime(new Date());
			company.setPersons(100 + i);
			company.setUpdateTime(Calendar.getInstance());
			companyDtoList.add(company);
			/*if (i % 30 == 0) {
				listWriter.process(companyDtoList, new String[]{"logo"});
				companyDtoList = new ArrayList<>();
			}*/
		}
		Stopwatch stopwatch = Stopwatch.createStarted();
		for(int i=1;i<=2;i++) {
			listWriter.process(companyDtoList);
		}

		OutputStream outputStream = new FileOutputStream("/Users/wangwei/test-async-large-list.xlsx");
		listWriter.export(outputStream);
		listWriter.close();
		log.info("export 2040000 records, execute time(seconds) :" + stopwatch.elapsed(TimeUnit.MILLISECONDS));
	}

	@Test
	public void testLargeList() throws Exception {
		// 11904 12086
		ExcelLargeListWriter listWriter = ExcelWriterFactory.createLargeListWriter("数据导出",100);
		listWriter.setNoneDataTips(true);
		List<CompanyDto2> companyDtoList = Lists.newArrayList();
		/*for (int i = 1; i <= 100000; i++) {
			CompanyDto2 company = new CompanyDto2();
			company.setAddress("address"+i);
			company.setName("短笛阿德南科技"+i);
			company.setCreateTime(new Date());
			company.setPersons(100+i);
			company.setUpdateTime(Calendar.getInstance());
			companyDtoList.add(company);
		}
		Stopwatch stopwatch = Stopwatch.createStarted();
		for(int i=1;i<=5;i++) {
			listWriter.process(companyDtoList,new String[]{"logo"});
		}*/
		OutputStream outputStream = new FileOutputStream("/Users/wangwei/test-large-list.xlsx");
		listWriter.export(outputStream);
		listWriter.close();
//		log.info("export 2040000 records, execute time(seconds) :"+stopwatch.elapsed(TimeUnit.MILLISECONDS));
	}

	/*public static void main(String[] args) throws Exception{
		Scanner scanner = new Scanner(System.in);
		String str = scanner.nextLine();
		System.out.println(str);

		String path = ExcelBootLoader.class.getClassLoader().getResource("excel-template").getPath();
		ExcelBootLoader.loadExcelTemplate(path);
		ExcelBootLoader.loadModel("io.terminus.model","io.terminus.model2");

		// 11904 12086
		ExcelLargeListWriter listWriter = ExcelWriterFactory.createLargeListWriter("数据导出",100);
		List<CompanyDto2> companyDtoList = Lists.newArrayList();
		InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();
		for (int i = 1; i <= 100000; i++) {
			CompanyDto2 company = new CompanyDto2();
			company.setAddress("address"+i);
			company.setName("短笛阿德南科技"+i);
			company.setCreateTime(new Date());
			company.setPersons(100+i);
			company.setUpdateTime(Calendar.getInstance());
			companyDtoList.add(company);
		}
		Stopwatch stopwatch = Stopwatch.createStarted();
		for(int i=1;i<=5;i++) {
			listWriter.process(companyDtoList,new String[]{"logo"});
		}
		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/test-large-list.xlsx");
		listWriter.export(outputStream);
		listWriter.close();
		log.info("export 2040000 records, execute time(seconds) :"+stopwatch.elapsed(TimeUnit.MILLISECONDS));
		str = scanner.nextLine();
		System.out.println(str);
	}*/

	@Test
	public void testAsyncExport() throws Exception {
		ExcelLargeListBatchWriter listWriter = ExcelWriterFactory.createLargeListBatchWriter("/Users/wangwei/Downloads/async");
		List<CompanyDto2> companyDtoList = Lists.newArrayList();
		for (int i = 1; i <= 2000; i++) {
			CompanyDto2 company = new CompanyDto2();
			company.setAddress("address"+i);
			company.setName("短笛阿德南科技"+i);
			company.setCreateTime(new Date());
			company.setPersons(100+i);
			company.setUpdateTime(Calendar.getInstance());
			companyDtoList.add(company);
		}
		Stopwatch stopwatch = Stopwatch.createStarted();
		for(int i=1;i<=2;i++) {
			listWriter.process(companyDtoList,"excel-"+i,"sheet",new String[]{"address"});
		}
		listWriter.export("excel");
		log.info("export 2040000 records, execute time(seconds) :"+stopwatch.elapsed(TimeUnit.SECONDS));
	}

	@Test
	public void testZip() throws Exception {
		File excelFile1 = new File("/Users/vachel/Downloads/wl-3.xlsx");
		File excelFile2 = new File("/Users/vachel/Downloads/wl-3.xls");
		File excelFile3 = new File("/Users/vachel/Downloads/sample-1-xls.xls");
		File excelFile4 = new File("/Users/vachel/Downloads/excel");
		File excelFile5 = new File("/Users/vachel/Downloads/Sweet-Excel-core-1.0-SNAPSHOT/test/wl-3.xls");
		List<File> fileList = Lists.newArrayList();
		fileList.add(excelFile1);
		fileList.add(excelFile2);
		fileList.add(excelFile3);
		fileList.add(excelFile4);

		fileList.add(excelFile5 );

		ZipCompressUtil compressUtil = new ZipCompressUtil();
		compressUtil.compressFile(fileList,"/Users/vachel/Downloads/excel.zip",null);

	}

	@Test
	public void test() throws Exception {
		File tempFile = File.createTempFile("aaa", "bbb",new File("/Users/vachel/Downloads/"));
		System.out.println(tempFile.getPath());
		System.out.println(tempFile.isDirectory());
	}
	@Test
	public void testWidth() throws Exception {
		Scanner scanner = new Scanner(System.in);
		String str = scanner.nextLine();


		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("sheet");
		for (int i = 0; i < 100000; i++) {
			XSSFRow row = sheet.createRow(i);
			for (int j = 0; j < 12; j++) {
				XSSFCell cell = row.createCell(j);
			}
		}
		Stopwatch stopwatch = Stopwatch.createStarted();
		for (int i = 0; i < 100000; i++) {
			for (int j = 0; j < 12; j++) {
				XSSFRow row = sheet.getRow(i);
				row.setHeightInPoints(50);
				sheet.setColumnWidth(ExcelConstant.ZERO_SHORT,50);
			}
		}
		log.info("export 2040000 records, execute time(seconds) :"+stopwatch.elapsed(TimeUnit.MILLISECONDS));
		while (true) {

		}

	}

	public static void main(String[] args){
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("sheet");
		for (int i = 0; i < 100000; i++) {
			XSSFRow row = sheet.createRow(i);
			for (int j = 0; j < 12; j++) {
				XSSFCell cell = row.createCell(j);
			}
		}
		Stopwatch stopwatch = Stopwatch.createStarted();
		for (int i = 0; i < 100000; i++) {
			for (int j = 0; j < 12; j++) {
				XSSFRow row = sheet.getRow(i);
				row.setHeightInPoints(50);
				sheet.setColumnWidth(ExcelConstant.ZERO_SHORT,50);
			}
		}
		log.info("export 2040000 records, execute time(seconds) :"+stopwatch.elapsed(TimeUnit.MILLISECONDS));
	}
}

