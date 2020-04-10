package com.github.excel.starter;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.AbstractImportTemplateExcluder;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import com.google.common.collect.Maps;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.write.ExcelLargeListWriter;
import com.github.excel.write.impl.ExcelLargeListWriterImpl;
import lombok.extern.java.Log;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.Resource;
import org.springframework.core.io.support.ResourcePatternResolver;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:18 下午
 * @Description: Excel加载自动配置
 */
@Configuration
@EnableConfigurationProperties(ExcelProperties.class)
@Log
public class ExcelLoadAutoConfiguration {

	@Autowired
	private ExcelProperties excelProperties ;

	@Autowired(required = false)
	private AbstractImportTemplateExcluder templateExcluder ;

	@Autowired
	private ResourcePatternResolver resourcePatternResolver;

	@Bean
	@ConditionalOnMissingBean
	public ExcelLargeListWriter excelWriter(){
		log.info("==========================");
		log.info("Powered by Sweet-Excel.");
		log.info("==========================");
		if(Objects.nonNull(excelProperties.getDtoPackage()) && !excelProperties.getDtoPackage().isEmpty()) {
			String[] packageArr = excelProperties.getDtoPackage().stream().toArray(String[]::new);
			ExcelBootLoader.loadModel(packageArr);
		}
		if (StringUtil.notEmpty(excelProperties.getExportTemplateDir())) {
			List<Map<String, Object>> fileList = new ArrayList<>();
			addResource2List("classpath*:" + excelProperties.getExportTemplateDir() + ExcelConstant.FILE_SEPARATOR + "*" + ExcelConstant.XLSX_STR, fileList);
			addResource2List("classpath*:" + excelProperties.getExportTemplateDir() + ExcelConstant.FILE_SEPARATOR + "*" + ExcelConstant.XLS_STR, fileList);
			ExcelBootLoader.loadExcelTemplate(fileList);
		}
		if (StringUtil.notEmpty(excelProperties.getImportTemplateDir())) {
			List<Map<String, Object>> fileList = new ArrayList<>();
			addResource2List("classpath*:" + excelProperties.getImportTemplateDir() + ExcelConstant.FILE_SEPARATOR + "*" + ExcelConstant.XLSX_STR, fileList);
			addResource2List("classpath*:" + excelProperties.getImportTemplateDir() + ExcelConstant.FILE_SEPARATOR + "*" + ExcelConstant.XLS_STR, fileList);
			ExcelBootLoader.loadImportExcelTemplate(fileList, templateExcluder);
		}

		ExcelBootLoader.cacheWorkbook();
		return new ExcelLargeListWriterImpl("sheet");
	}

	private void addResource2List(String path, List<Map<String, Object>> files){
		try {
			log.info("load resource from " + path);
			Resource[] resources = resourcePatternResolver.getResources(path);
			if(resources != null){
				for(Resource resource: resources){
					Map<String, Object> map = Maps.newHashMap();
					map.put("name", resource.getFilename());
					log.info("load excel template with name " + resource.getFilename());
					map.put("input", resource.getInputStream());
					map.put("cacheInput", resource.getInputStream());
					files.add(map);
				}
			}
		} catch (IOException e){
			log.info("load resources error with Sweet-Excel, cause:" + Throwables.getStackTraceAsString(e));
			throw new ExcelReadException("template.load.fail");
		}
	}
}
