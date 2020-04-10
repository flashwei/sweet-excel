package com.github;

import com.github.excel.boot.ExcelBootLoader;
import com.github.read.template.ImportTemplateExcluder;
import lombok.extern.slf4j.Slf4j;
import org.junit.BeforeClass;


/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
@Slf4j
public class BaseExcelTest {
	@BeforeClass
	public static void loadModel(){
		String path = ExcelBootLoader.class.getClassLoader().getResource("excel-template").getPath();
		String path1 = ExcelBootLoader.class.getClassLoader().getResource("import-excel-template").getPath();
		ExcelBootLoader.loadExcelTemplate(path);
		ImportTemplateExcluder importTemplateExcluder = new ImportTemplateExcluder();
		importTemplateExcluder.addTemplateExclude();

		ExcelBootLoader.loadImportExcelTemplate(path1,importTemplateExcluder);
		ExcelBootLoader.loadModel("com.github.model","com.github.model2");
		ExcelBootLoader.cacheWorkbook();

	}
}

