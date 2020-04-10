package com.github.read;

import com.github.BaseExcelTest;
import com.github.model.ProjectBidsBean;
import com.github.model.UserExcelDtoImportBean;
import com.google.common.base.Stopwatch;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.model.ExcelReadErrorMsgModel;
import com.github.excel.read.ExcelReadBatchProcess;
import com.github.excel.read.ExcelReader;
import com.github.excel.read.ExcelReaderFactory;
import com.github.excel.read.ExcelUserReader;
import com.github.model.UserExcelDtoImportBean1;
import com.github.model.UserExcelDtoImportList;
import com.github.model.UserExcelDtoImportList1;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.TimeUnit;


/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
@Slf4j
public class ExcelTestRead extends BaseExcelTest {

	@Test
	public void testSample() throws Exception {

		InputStream is = new FileInputStream("/Users/vachel/Downloads/sample-import.xlsx");
		ExcelReader reader = ExcelReaderFactory.createUserReader("sample-import.xlsx",is);
		reader.setReadPicture(true);
		reader.addModel(UserExcelDtoImportBean.class,0);
		reader.addModel(UserExcelDtoImportBean1.class,0);
//		reader.addModel(UserExcelDtoImportBean2.class,1);

//		reader.addModelList(UserExcelDtoImportList.class, 0);
		reader.addModelList(UserExcelDtoImportList1.class,0, new ExcelReadBatchProcess<UserExcelDtoImportList1>() {
			@Override
			public int getBatchSize() {
				return 3;
			}

			@Override
			public void process(List<UserExcelDtoImportList1> dataList) {
//				throw new ExcelReadException("heh");
				log.info("UserExcelDtoImportList1 ==============");
				dataList.forEach(e-> System.out.println(e));
				log.info("==============");
			}
		});

		Stopwatch stopwatch = Stopwatch.createStarted();
		ExcelReadErrorMsgModel errorMsgModel = reader.parseWithError();
//		reader.parse();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "时间");
		/*log.info("parse exists error: "+errorMsgModel.getExistsError());
		errorMsgModel.getErrorMsgInfoList().forEach(e ->{
			log.error(e.getErrorMsg());
		});*/
		/*UserExcelDtoImportBean model = reader.getModel(UserExcelDtoImportBean.class);
		System.out.println(model);
		UserExcelDtoImportBean1 model1 = reader.getModel(UserExcelDtoImportBean1.class);
		System.out.println(model1);
		UserExcelDtoImportBean2 model2 = reader.getModel(UserExcelDtoImportBean2.class);
		System.out.println(model2);
		List<UserExcelDtoImportList> modelList= reader.getModelList(UserExcelDtoImportList.class);
		List<UserExcelDtoImportList1> modelList1= reader.getModelList(UserExcelDtoImportList1.class);
		modelList.forEach(e -> System.out.println(e));*/
//		modelList1.forEach(e -> System.out.println(e));
		/*int i=1;
		for(ReadPictureModel logo:modelList1.get(4).getLogo()) {
			OutputStream out = new FileOutputStream("/Users/vachel/Downloads/test-logo"+i+ logo.getSuffix());
			out.write(logo.getBytes());
			out.close();
			++ i ;
		}*/
	}

	@Test
	public void testReadList() throws Exception {

		InputStream is = new FileInputStream("/Users/vachel/Downloads/import-test.xlsx");
		ExcelReader reader = ExcelReaderFactory.createUserReader("import-test.xlsx",is);
		reader.setReadPicture(true);
//		reader.addModel(UserExcelDtoImportBean.class,0);
		reader.addModelList(UserExcelDtoImportList.class,0);
		reader.addModelList(UserExcelDtoImportList1.class,0);
//		reader.addModel(UserExcelDtoImportBean1.class,0);
//		reader.addModel(UserExcelDtoImportBean2.class,1);
//		reader.addModelList(UserExcelDtoImportList1.class,1);
		Stopwatch stopwatch = Stopwatch.createStarted();
		reader.parse();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "时间");
		/*log.info("parse exists error: "+errorMsgModel.getExistsError());
		errorMsgModel.getErrorMsgInfoList().forEach(e ->{
			log.error(e.getErrorMsg());
		});*/
		List<UserExcelDtoImportList> modelList= reader.getModelList(UserExcelDtoImportList.class);
		List<UserExcelDtoImportList1> modelList1= reader.getModelList(UserExcelDtoImportList1.class);
		modelList.forEach(e -> System.out.println(e));
		System.out.println("==============");
		modelList1.forEach(e -> System.out.println(e));
	}
	@Test
	public void testReadProject() throws Exception {

		InputStream is = new FileInputStream("/Users/vachel/Documents/Sweet-Excel/导入模板/project-bids.xlsx");
		ExcelReader reader = ExcelReaderFactory.createUserReader("project-bids.xlsx",is,"project-bids.xlsx");
		reader.addModelList(ProjectBidsBean.class,1);
		Stopwatch stopwatch = Stopwatch.createStarted();
		ExcelReadErrorMsgModel errorMsgModel = reader.parseWithError();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "时间");
		log.info("parse exists error: "+errorMsgModel.getExistsError());
		errorMsgModel.getErrorMsgInfoList().forEach(e ->{
			log.error(e.getErrorMsg());
		});
		List<ProjectBidsBean> modelList = reader.getModelList(ProjectBidsBean.class);
		for (ProjectBidsBean bidsBean : modelList) {
			System.out.println(bidsBean);
		}
	}

	@Test
	public void testRead() throws Exception {
		InputStream is = new FileInputStream("/Users/vachel/Downloads/sample-1.xlsx");
		ExcelReader reader = new ExcelUserReader("sample-1.xlsx",is);
		reader.addModel(UserExcelDtoImportBean.class,0);
		reader.addModelList(UserExcelDtoImportList.class,0);
		Stopwatch stopwatch = Stopwatch.createStarted();
		reader.parse();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
		UserExcelDtoImportBean model = reader.getModel(UserExcelDtoImportBean.class);
		System.out.println(model);
	}

	@Test
	public void testExport() throws Exception {
		byte[] bytes = ExcelBootLoader.getExcelImportTemplateFileCacheMapValue("project-bids.xlsx");
		if (Objects.nonNull(bytes) && bytes.length > ExcelConstant.ZERO_SHORT) {
			OutputStream outputStream = new FileOutputStream(new File("/Users/vachel/Downloads/test11111.xlsx"));
			outputStream.write(bytes);
			outputStream.flush();
			outputStream.close();
		}

	}
	@Test
	public void testNumberRegex() throws Exception {
		boolean flag = ".2341".matches(ExcelConstant.NUMBER_PATTERN);
		System.out.println(flag);
		//43389.64541107639
		//1540199621691
		System.out.println(System.currentTimeMillis());
	}
}
