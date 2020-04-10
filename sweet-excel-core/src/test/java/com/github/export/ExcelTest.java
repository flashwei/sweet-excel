package com.github.export;

import com.google.common.base.Stopwatch;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.github.BaseExcelTest;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ComboBoxModel;
import com.github.excel.model.CommentModel;
import com.github.excel.model.ExcelCustomColumnModel;
import com.github.excel.model.ExcelMergeCustomColumnModel;
import com.github.excel.model.NumberScopeModel;
import com.github.excel.write.ExcelBasicStyle;
import com.github.excel.write.ExcelCustomWriter;
import com.github.excel.write.ExcelWriter;
import com.github.excel.write.ExcelWriterFactory;
import com.github.model.CompanyDto;
import com.github.model.MatterDto4;
import com.github.model.UserExcelDto;
import com.github.model.UserExcelDto2;
import com.github.model.UserExcelDto3;
import com.github.model2.CompanyDto2;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;


/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:26 下午
 * @Description: excel 测试
 */
@Slf4j
public class ExcelTest extends BaseExcelTest {


	@Test
	public void test() throws Exception {
		Workbook workbook = new XSSFWorkbook();

		Sheet sheet = workbook.createSheet("test");
		Row row = sheet.createRow(1);
		Cell cell = row.createCell(1);
		cell.setCellValue("aasddfdfd");
		CellStyle style = workbook.createCellStyle();

		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLUE.getIndex());

		cell.setCellStyle(style);

		style.setBorderTop(BorderStyle.MEDIUM_DASHED);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());

		CreationHelper createHelper = workbook.getCreationHelper();
		style.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss.SSS"));

		cell = row.createCell(2);
		cell.setCellValue("hello");

		workbook.write(new FileOutputStream("/Users/vachel/dev-soft/excel/test.xlsx"));
		workbook.close();
		log.info("导出excel完成");

		ExcelCustomWriter customWrite = (wb) -> {

		};
	}

	@Test
	public void testLoad() throws Exception {
		ExcelBootLoader boot = new ExcelBootLoader();
		boot.loadModel("io.terminus.model");
		Integer integer = new Integer(1);
		BigDecimal decimal = new BigDecimal(22);
		System.out.println(decimal instanceof Number);
	}


	@Test
	public void testSampleList() throws Exception {
		InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();
		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();
		List<UserExcelDto3> userList = new ArrayList<>();
		UserExcelDto3 userExcelDto2 = new UserExcelDto3();
		CommentModel commentModel = new CommentModel();
		commentModel.setCommentFontName(ExcelBasicStyle.FONT_SIZE16_BLOLD_RED);
		commentModel.setCommentText("这是测试comment");
		commentModel.setValue("多点点");
		userExcelDto2.setSexStr(commentModel);
		userExcelDto2.setName("terminus");
		userExcelDto2.setAge(5);
		userExcelDto2.setHeight(178.221f);
		userExcelDto2.setSex((byte) 1);
		userExcelDto2.setNickName("端点科技");
		userExcelDto2.setAvater("http://pmp.terminus.io/images/logo_reverse.png");
//			userExcelDto2.setLogo(bytes);
		userExcelDto2.setCreateTime(new Date());
		for (int i = 1; i <= 1000; i++) {
			userList.add(userExcelDto2);
		}
		NumberScopeModel scopeModel = new NumberScopeModel();
		scopeModel.setStart("1");
		scopeModel.setEnd("2");
		scopeModel.setCommentText("性别正确填写");
		ComboBoxModel comboBoxModel = new ComboBoxModel();
		comboBoxModel.setOptions(new String[]{"5", "10", "15", "20"});
		comboBoxModel.setCommentText("年龄要正确填写！");
		comboBoxModel.setCommentFontName(ExcelBasicStyle.FONT_SIZE16_BLOLD_RED);

		writer.addModelList(userList, "test", false).excludes(new String[]{"name"}).addValidationOrComment("sex", scopeModel).addValidationOrComment("age", comboBoxModel);
		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/sample-2.xlsx");
		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "sample-2", ExcelSuffixEnum.XLSX);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testExport() throws Exception {
//		InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
//		byte[] bytes = IOUtils.toByteArray(is);
//		is.close();

		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();

		UserExcelDto userExcelDto2 = new UserExcelDto();
		userExcelDto2.setName("terminus");
		userExcelDto2.setAge(20);
		userExcelDto2.setAgeShort((short) 21);
		userExcelDto2.setHeight(178.221f);
		userExcelDto2.setHeightDouble(178.1333);
		userExcelDto2.setLock(false);
		userExcelDto2.setMoney(1000000L);
		userExcelDto2.setSex((byte) 1);
		userExcelDto2.setMoneyBig(new BigDecimal("10000212"));
		userExcelDto2.setCreateTime(new Date());
		userExcelDto2.setUpdateTime(Calendar.getInstance());
		userExcelDto2.setNickName("端点科技");
		userExcelDto2.setAvater("http://pmp.terminus.io/images/logo_reverse.png");
		userExcelDto2.setEmail("mailto:wangwei@terminus.io");
//		userExcelDto2.setLogo(bytes);

		writer.addModel(userExcelDto2, "测试", true);

		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/test01.xlsx");
		/*ExcelCustomWriter customWriter = (wb) -> {
			Sheet sheet = wb.createSheet("测试-1");
			Row row = sheet.createRow(2);
			Cell cell = row.createCell(3);
			cell.setCellStyle(writer.getStyle(ExcelBasicStyle.STYLE_HLINK));
			cell.setCellValue("填充字符串");
		};

		writer.setCustomWrite(customWriter);*/
		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "test-export", ExcelSuffixEnum.XLSX);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testExportTemplate() throws Exception {

		InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();

		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();
		writer.setStreaming(true);

		/*UserExcelDto userExcelDto2 = new UserExcelDto();
		userExcelDto2.setName("terminus");
		userExcelDto2.setAge(20);
		userExcelDto2.setAgeShort((short) 21);
		userExcelDto2.setHeight(178.221f);
		userExcelDto2.setHeightDouble(178.1333);
		userExcelDto2.setLock(false);
		userExcelDto2.setMoney(1000000L);
		userExcelDto2.setSex((byte) 1);
		userExcelDto2.setMoneyBig(new BigDecimal("10000212"));
		userExcelDto2.setCreateTime(new Date());
		userExcelDto2.setUpdateTime(Calendar.getInstance());
		userExcelDto2.setNickName("端点科技");
		userExcelDto2.setAvater("http://pmp.terminus.io/images/logo_reverse.png");
		userExcelDto2.setEmail("mailto:wangwei@terminus.io");
		userExcelDto2.setLogo(bytes);

		CompanyDto company1 = new CompanyDto();
		company1.setAddress("address");
		company1.setName("短笛阿德南科技");
		company1.setCreateTime(new Date());
		company1.setPersons(100);
		company1.setLogo(bytes);
		company1.setUpdateTime(Calendar.getInstance());
		userExcelDto2.setCompany(company1);*/


		List<CompanyDto> companyDtoList = Lists.newArrayList();

		for (int i = 1; i <= 500000; i++) {
			CompanyDto company = new CompanyDto();
			company.setAddress("address" + i);
			company.setName("短笛阿德南科技" + i);
			company.setCreateTime(new Date());
			company.setPersons(100 + i);
//			company.setLogo(bytes);
			company.setUpdateTime(Calendar.getInstance());
			companyDtoList.add(company);
		}
		writer.addModelList(companyDtoList, "测试", true);
		/*writer.addModelList(companyDtoList,"测试",true);
		ExcelCustomColumnModel columnModel = new ExcelCustomColumnModel();
		columnModel.setColIndex(2);
		columnModel.setRowIndex(2);
		columnModel.setSheetName("test");
		columnModel.setValue("自定义导出");
		columnModel.setStyleName(ExcelBasicStyle.STYLE_TITLE_RED_FONT);
		columnModel.setColWidth((short) 200);
		columnModel.setRowHeight((short)50);
		writer.addCustomColumn(columnModel);

		ExcelMergeCustomColumnModel mergeCustomColumnModel = new ExcelMergeCustomColumnModel();
		mergeCustomColumnModel.setFirstColumn(6);
		mergeCustomColumnModel.setFirstRow(6);
		mergeCustomColumnModel.setLastColumn(7);
		mergeCustomColumnModel.setLastRow(7);
		mergeCustomColumnModel.setSheetName("test");
		mergeCustomColumnModel.setValue("合并单元格");
		mergeCustomColumnModel.setStyleName(ExcelBasicStyle.STYLE_TITLE_RED_FONT);
		mergeCustomColumnModel.setColWidth((short) 200);
		mergeCustomColumnModel.setRowHeight((short)50);
		writer.addMergeCustomColumn(mergeCustomColumnModel);*/
		writer.setStreaming(true);
		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/test-export-template.xlsx");

		/*ExcelCustomWriter customWriter = (wb) -> {
			Sheet sheet = wb.createSheet("测试-1");
			Row row = sheet.createRow(2);
			Cell cell = row.createCell(3);
			cell.setCellStyle(writer.getStyle(ExcelBasicStyle.STYLE_HLINK));
			cell.setCellValue("填充字符串");
		};*/

//		writer.setCustomWrite(customWriter);
		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "test-export", ExcelSuffixEnum.XLSX);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testModelPackage() throws Exception {

		/*InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();*/

		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();
		writer.addStyle(ExcelCustomStyle.class);

		List<CompanyDto2> companyDtoList = Lists.newArrayList();

		for (int i = 1; i <= 500; i++) {
			CompanyDto2 company = new CompanyDto2();
			company.setAddress("address" + i);
			company.setName("短笛阿德南科技" + i);
			company.setCreateTime(new Date());
			company.setPersons(100 + i);
//			company.setLogo(bytes);
			company.setUpdateTime(Calendar.getInstance());
			companyDtoList.add(company);
		}
		writer.addModelList(companyDtoList, "测试", true);
		writer.setStreaming(true);
		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/test-export-template.xls");

		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "test-export", ExcelSuffixEnum.XLS);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testTemplate() throws Exception {

		InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();

		ExcelWriter writer = ExcelWriterFactory.createUserModelWriterWithTemplate("test1.xls");

		UserExcelDto userExcelDto2 = new UserExcelDto();
		userExcelDto2.setName("terminus");
		userExcelDto2.setAge(20);
		userExcelDto2.setAgeShort((short) 21);
		userExcelDto2.setHeight(178.221f);
		userExcelDto2.setHeightDouble(178.1333);
		userExcelDto2.setLock(false);
		userExcelDto2.setMoney(1000000L);
		userExcelDto2.setSex((byte) 1);
		userExcelDto2.setMoneyBig(new BigDecimal("10000212"));
		userExcelDto2.setCreateTime(new Date());
		userExcelDto2.setUpdateTime(Calendar.getInstance());
		userExcelDto2.setNickName("端点科技");
		userExcelDto2.setAvater("http://pmp.terminus.io/images/logo_reverse.png");
		userExcelDto2.setEmail("mailto:wangwei@terminus.io");
		userExcelDto2.setLogo(bytes);

		CompanyDto company1 = new CompanyDto();
		company1.setAddress("address");
		company1.setName("短笛阿德南科技");
		company1.setCreateTime(new Date());
		company1.setPersons(100);
		company1.setLogo(bytes);
		company1.setUpdateTime(Calendar.getInstance());
		userExcelDto2.setCompany(company1);


		List<CompanyDto> companyDtoList = Lists.newArrayList();

		for (int i = 1; i <= 100; i++) {
			CompanyDto company = new CompanyDto();
			company.setAddress("address" + i);
			company.setName("短笛阿德南科技" + i);
			company.setCreateTime(new Date());
			company.setPersons(100 + i);
//			company.setLogo(bytes);
			company.setUpdateTime(Calendar.getInstance());
			companyDtoList.add(company);
		}
		writer.addModelList(companyDtoList, "测试", true);
		writer.addModelList(companyDtoList, "测试", true);
		ExcelCustomColumnModel columnModel = new ExcelCustomColumnModel();
		columnModel.setColIndex(2);
		columnModel.setRowIndex(2);
		columnModel.setSheetName("test");
		columnModel.setValue("自定义导出");
		columnModel.setStyleName(ExcelBasicStyle.STYLE_TITLE_RED_FONT);
		columnModel.setColWidth((short) 200);
		columnModel.setRowHeight((short) 50);
		writer.addCustomColumn(columnModel);

		ExcelMergeCustomColumnModel mergeCustomColumnModel = new ExcelMergeCustomColumnModel();
		mergeCustomColumnModel.setFirstColumn(6);
		mergeCustomColumnModel.setFirstRow(6);
		mergeCustomColumnModel.setLastColumn(7);
		mergeCustomColumnModel.setLastRow(7);
		mergeCustomColumnModel.setSheetName("test");
		mergeCustomColumnModel.setValue("合并单元格");
		mergeCustomColumnModel.setStyleName(ExcelBasicStyle.STYLE_TITLE_RED_FONT);
		mergeCustomColumnModel.setColWidth((short) 200);
		mergeCustomColumnModel.setRowHeight((short) 50);
		writer.addMergeCustomColumn(mergeCustomColumnModel);
		writer.setStreaming(true);
		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/test-template.xls");

		ExcelCustomWriter customWriter = (wb) -> {
			Sheet sheet = wb.getSheet("测试");
			ExcelHelper.createComment(wb, sheet, 0, 0, "vachel", "姓名要正确填写哟！", writer.getFont(ExcelBasicStyle.FONT_SIZE16_BLOLD_RED));
		};

		writer.setCustomWrite(customWriter);
		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "test-export", ExcelSuffixEnum.XLS);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testExportList() throws Exception {
		InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();

		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();

		UserExcelDto2 userExcelDto2 = new UserExcelDto2();
		userExcelDto2.setName("terminus");
		userExcelDto2.setAge(20);
		userExcelDto2.setAgeShort((short) 21);
		userExcelDto2.setHeight(178.221f);
		userExcelDto2.setHeightDouble(178.1333);
		userExcelDto2.setLock(false);
		userExcelDto2.setMoney(1000000L);
		userExcelDto2.setSex((byte) 1);
		userExcelDto2.setMoneyBig(new BigDecimal("10000212"));
		userExcelDto2.setCreateTime(new Date());
		userExcelDto2.setUpdateTime(Calendar.getInstance());
		userExcelDto2.setNickName("端点科技");
		userExcelDto2.setAvater("http://pmp.terminus.io/images/logo_reverse.png");
		userExcelDto2.setEmail("mailto:wangwei@terminus.io");
		userExcelDto2.setLogo(bytes);
		Map<String, String> map = Maps.newHashMap();
		map.put("aaaa", "111");
		map.put("bbb", "www");
		map.put("ccc", "333");
		userExcelDto2.setMap(map);

//		writer.addModel(userExcelDto2, "测试",true);

		List<UserExcelDto> userlist = Lists.newArrayList();
		for (int i = 1; i <= 10; i++) {
			UserExcelDto userExcelDto = new UserExcelDto();
			userExcelDto.setName("terminus" + i);
			userExcelDto.setAge(20 + i);
			userExcelDto.setAgeShort((short) (21 + i));
			userExcelDto.setHeight(178.221f + i);
			userExcelDto.setHeightDouble(178.1333 + i);
			userExcelDto.setLock(false);
			userExcelDto.setMoney(1000000L + i);
			userExcelDto.setSex((byte) 1);
			userExcelDto.setMoneyBig(new BigDecimal("10000212" + i));
			userExcelDto.setCreateTime(new Date());
			userExcelDto.setUpdateTime(Calendar.getInstance());
			userExcelDto.setNickName("端点科技" + i);
			userExcelDto.setAvater("http://pmp.terminus.io/images/logo_reverse.png");
			userExcelDto.setEmail("mailto:wangwei@terminus.io" + i);
			userExcelDto.setLogo(bytes);

			userlist.add(userExcelDto);
		}

		writer.addModelList(userlist, "测试", false);

		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/test-export-list.xlsx");

		ExcelCustomWriter customWriter = (wb) -> {
			Sheet sheet = wb.createSheet("测试-1");
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
//			cell.setCellStyle(writer.getStyle(ExcelBasicStyle.STYLE_HLINK));
			cell.setCellValue(1);
			cell = row.createCell(1);
			cell.setCellValue(2);
			cell = row.createCell(2);
			cell.setCellFormula("sum(A1:B1)");


			ExcelHelper.addRefersToFormula(wb, "name1", "SUM('测试-1'!$A1:$B1)");

			cell = row.createCell(4);
			cell.setCellFormula("name1");
			log.info("rownum : " + ExcelHelper.getRowNum(wb, "测试"));
			log.info("colNum : " + ExcelHelper.getCellNum(wb, "测试", 8));
			ExcelHelper.addDropDownValidation(sheet, new CellRangeAddressList(0, 3, 0, 3), new String[]{"aaa", "bbb", "ccc"});
			ExcelHelper.addRangeValidation(sheet, new CellRangeAddressList(4, 6, 4, 6), "10", "100");


		};
		ExcelCustomWriter customWriter1 = (wb) -> {
			Sheet sheet = wb.createSheet("测试-1");
			ExcelHelper.addRangeValidation(sheet, new CellRangeAddressList(4, 6, 4, 6), "10", "100");
		};

		writer.setCustomWrite(customWriter);
		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "test-export-list", ExcelSuffixEnum.XLSX);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info("total seconds:" + elapsed);

	}

	@Test
	public void testSample() throws Exception {

		InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();

		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();

		UserExcelDto3 userExcelDto2 = new UserExcelDto3();
		userExcelDto2.setName("terminus");
		userExcelDto2.setAge(5);
		userExcelDto2.setHeight(178.221f);
//		userExcelDto2.setSex((byte) 1);
		userExcelDto2.setNickName("端点科技");
		userExcelDto2.setAvater("http://pmp.terminus.io/images/logo_reverse.png");
		userExcelDto2.setLogo(bytes);
		CommentModel commentModel = new CommentModel();
		commentModel.setCommentFontName(ExcelBasicStyle.FONT_SIZE16_BLOLD_RED);
		commentModel.setCommentText("这是测试comment");
		commentModel.setValue("多点点");
//		userExcelDto2.setCompanyType("央企");
		ComboBoxModel boxModel = new ComboBoxModel();
		boxModel.setOptions(new String[]{"企业1", "企业2"});
		boxModel.setCommentText("好好填写");
		boxModel.setValue("123");
		userExcelDto2.setCompanyType1(boxModel);

		NumberScopeModel scopeModel = new NumberScopeModel();
		scopeModel.setEnd("100");
		scopeModel.setStart("0");
		scopeModel.setCommentText("年龄范围要好好填写！");


		userExcelDto2.setScopeModel(scopeModel);

		userExcelDto2.setSexStr(commentModel);
		userExcelDto2.setCreateTime(new Date());

		writer.addModel(5, 5, userExcelDto2, "test", false);
		writer.addStyle(ExcelCustomStyle.class);
		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/sample-1.xls");
		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "sample-1", ExcelSuffixEnum.XLS);
		outputStream.flush();
		outputStream.close();

		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testTable() throws Exception {
		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();
		writer.setListCla(MatterDto4.class);
		writer.addStyle(ExcelCustomStyle.class);
		/*writer.setStreaming(true);
		MatterDto4 matterDto4 = new MatterDto4();
		matterDto4.setMatterName("桌子");
		matterDto4.setMatterCode("CODE-123456");
		matterDto4.setBrand("永久牌");
		writer.addModel(matterDto4, "测试",false);*/
		OutputStream outputStream = new FileOutputStream("/Users/vachel/Downloads/test-table.xlsx");
		writer.process(outputStream, "test-table", ExcelSuffixEnum.XLSX);

	}


}

