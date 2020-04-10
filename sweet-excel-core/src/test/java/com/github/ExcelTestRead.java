package com.github;

import com.github.excel.constant.ExcelConstant;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;


/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
@Slf4j
public class ExcelTestRead {

	@Test
	public void testSample() throws Exception {
//		ExcelLoadModelBoot boot = new ExcelLoadModelBoot();
//		boot.loadModel("io.terminus.model");
//
//		InputStream is = new FileInputStream("/Users/vachel/Downloads/sample-1.xlsx");
//		ExcelReader reader = new ExcelUserReader<>("sample-1.xlsx",is);
//		reader.addModel(UserExcelDto3Import.class,0);
//		reader.addModelList(UserExcelDto4Import.class,0);
//		Stopwatch stopwatch = Stopwatch.createStarted();
//		ExcelReadErrorMsgModel errorMsgModel = reader.parseWithError();
//		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
//		log.info(elapsed + "时间");
//		log.info("parse exists error: "+errorMsgModel.getExistsError());
//		errorMsgModel.getErrorMsgInfoList().forEach(e ->{
//			log.error(e.getErrorMsg());
//		});
//		UserExcelDto3Import model = reader.getModel(UserExcelDto3Import.class);
//		System.out.println(model);
//		List<UserExcelDto4Import> modelList= reader.getModelList(UserExcelDto4Import.class);
//		modelList.forEach(e -> System.out.println(e));
//	}
//	@Test
//	public void testRead() throws Exception {
//		ExcelLoadModelBoot boot = new ExcelLoadModelBoot();
//		boot.loadModel("io.terminus.model");
//
//		InputStream is = new FileInputStream("/Users/vachel/Downloads/sample-1.xlsx");
//		ExcelReader reader = new ExcelUserReader<>("sample-1.xlsx",is);
//		reader.addModel(UserExcelDto3Import.class,0);
//		reader.addModel(UserExcelDto4Import.class,0);
//		Stopwatch stopwatch = Stopwatch.createStarted();
//		reader.parse();
//		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
//		log.info(elapsed + "");
//		UserExcelDto3Import model = reader.getModel(UserExcelDto3Import.class);
//		System.out.println(model);
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
