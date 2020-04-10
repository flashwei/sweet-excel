package com.github.export;

import com.github.excel.write.*;
import com.google.common.base.Stopwatch;
import com.google.common.collect.Lists;
import com.github.BaseExcelTest;
import com.github.excel.boot.WorkbookCachePool;
import com.github.excel.enums.ExcelExportColumnFillTypeEnum;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ExcelCustomColumnModel;
import com.github.excel.model.ExcelMergeCustomColumnModel;
import com.github.model.CompanyDto;
import com.github.model.MatterDto;
import com.github.model.MatterDto2;
import com.github.model.OfferDto;
import com.github.model.UserExcelDto;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;



@Slf4j
public class MatterTest  extends BaseExcelTest {

	@Test
	public void testWl2() throws Exception {

		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();
		writer.setStreaming(true);
		List<MatterDto> matterDtos = Lists.newArrayList();

		for (int i = 1; i <= 500; i++) {
			MatterDto dto = MatterDto.builder().matterCode("120042981"+i).matterName("玻璃/钢化5*480*1528"+i).brand("apple"+i)
					.weight("weight"+i).unit("PC"+i).supplier("测试网络信息技术有限公司"+i).prevPurchasePrice(11.25+i).initialPrice(10.0+i)
					.finalPrice(22.1+i).rate(0.33+i).deliveryDate(7+i).freight("供方承担"+i).payType("现金"+i).concatMobile("13916896350"+i)
					.concatName("杜智诚"+i).suggest("建议采购"+i).remark("备注"+i).queryTime(new Date()).build();
			matterDtos.add(dto);
		}

		writer.addModelList(matterDtos,"直接比价",true);


		File file = new File("/Users/wangwei/Downloads/wl-2.xlsx");
		OutputStream outputStream = new FileOutputStream(file);

//		OutputStream outputStream = new FileOutputStream("/Users/vachel/dev-soft/excel/wl-2.xls");

		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream,"wl-2",ExcelSuffixEnum.XLSX);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testWlLarge() throws Exception {

		ExcelLargeListWriter writer = ExcelWriterFactory.createLargeListMultiThreadWriter("ss", 1000);
		List<MatterDto> matterDtos = Lists.newArrayList();

		for (int i = 1; i <= 1000; i++) {
			MatterDto dto = MatterDto.builder().matterCode("120042981"+i).matterName("玻璃/钢化5*480*1528"+i).brand("apple"+i)
					.weight("weight"+i).unit("PC"+i).supplier("测试网络信息技术有限公司"+i).prevPurchasePrice(11.25+i).initialPrice(10.0+i)
					.finalPrice(22.1+i).rate(0.33+i).deliveryDate(7+i).freight("供方承担"+i).payType("现金"+i).concatMobile("13916896350"+i)
					.concatName("杜智诚"+i).suggest("建议采购"+i).remark("备注"+i).queryTime(new Date()).build();
			matterDtos.add(dto);
		}

		writer.process(matterDtos);
		writer.process(matterDtos);


		File file = new File("/Users/wangwei/Downloads/wl-large.xlsx");
		OutputStream outputStream = new FileOutputStream(file);

//		OutputStream outputStream = new FileOutputStream("/Users/vachel/dev-soft/excel/wl-2.xls");

		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.export(outputStream);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}

	@Test
	public void testWl3() throws Exception {
		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();

		List<MatterDto2> matterDtos = Lists.newArrayList();

		for (int i = 1; i <= 10; i++) {
			MatterDto2 dto = MatterDto2.builder().matterCode("120042981"+i).matterName("玻璃/钢化5*480*1528"+i).brand("apple"+i).num(100+i)
					.unit("PC"+i).build();
			matterDtos.add(dto);
		}

		writer.addModelList(matterDtos,"测试",false);
		List<OfferDto> offerDtos = Lists.newArrayList();

		for (int i = 1; i <= 10; i++) {
			OfferDto dto = OfferDto.builder().cleanPrice(10.9+i).finalPrice(20.91+i).initPrice(8.8+i).price(22.2+i).build();
			offerDtos.add(dto);
		}
		int listColIndex = 5 ;

		for(int i=1;i<=3;i++) {
			writer.addModelList(1, listColIndex, offerDtos, "测试", false);
			listColIndex+=4;
		}

		int colIndex = 5 ;
		for(int i=1;i<=3;i++) {
			ExcelMergeCustomColumnModel columnModel = new ExcelMergeCustomColumnModel();
			columnModel.setFirstColumn(colIndex); // 开始列，下标0对应第一列 开始
			columnModel.setFirstRow(0);// 开始行，小标0对应第一行
			columnModel.setLastColumn(colIndex+3); // 结束列
			columnModel.setLastRow(0);// 结束行
			columnModel.setSheetName("测试"); // shetName
			columnModel.setValue("徐志彪"+i); // 设置value
			columnModel.setStyleName(ExcelBasicStyle.STYLE_TITLE_RED_FONT); // 设置样式
			writer.addMergeCustomColumn(columnModel); // 添加合并
			colIndex+=4;
		}

		MatterDto dto = MatterDto.builder().matterCode("120042981").matterName("玻璃/钢化5*480*1528").brand("apple")
				.weight("weight").unit("PC").supplier("测试网络信息技术有限公司").prevPurchasePrice(11.25).initialPrice(10.0)
				.finalPrice(22.1).rate(0.33).deliveryDate(7).freight("供方承担").payType("现金").concatMobile("13916896350")
				.concatName("杜智诚").suggest("建议采购").queryTime(new Date()).build();
		writer.addModel(dto,"测试",true).excludes(new String[]{"concatMobile"});

		/*InputStream is = new FileInputStream("/Users/vachel/Downloads/aaa.png");
		byte[] bytes = IOUtils.toByteArray(is);
		is.close();*/

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

		CompanyDto company1 = new CompanyDto();
		company1.setAddress("address");
		company1.setName("杭州端点网络科技");
		company1.setCreateTime(new Date());
		company1.setPersons(100);
//		company1.setLogo(bytes);
		company1.setUpdateTime(Calendar.getInstance());

		UserExcelDto createUser = new UserExcelDto();
		createUser.setName("端点创始人");
		company1.setCreateUser(createUser);
		userExcelDto2.setCompany(company1);

		writer.addModel(userExcelDto2,"测试",true).excludes(new String[]{"company"});

		int rowIndex = 16;
		for(int i=1;i<=3;i++){
			ExcelCustomColumnModel model = new ExcelCustomColumnModel();
			model.setValue("审批人"+i);
			model.setRowIndex(rowIndex);
			model.setColIndex(0);
			model.setFillType(ExcelExportColumnFillTypeEnum.APPEND);
			model.setSheetName("测试");
			rowIndex++;
			writer.addCustomColumn(model);
		}

		ExcelCustomWriter customWriter = (wb) -> {
			Sheet sheet = wb.getSheet("测试");
			ExcelHelper.createComment(wb,sheet,0,0,"vachel","姓名要正确填写哟！",null);
			ExcelHelper.setSheetZoom(wb, "test", 200);
			ExcelHelper.setFooterNumberByDefault(wb,"test");
			ExcelHelper.setPrintArea(wb,1,0,10,0,15);
		};
		writer.selectSheet("test");
		writer.setCustomWrite(customWriter);
		File file = new File("/Users/wangwei/Downloads/wl-3.xlsx");
		OutputStream outputStream = new FileOutputStream(file);

		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream,"wl-3",ExcelSuffixEnum.XLSX);
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
	}
	@Test
	public void testList() throws Exception {
		ExcelWriter writer = ExcelWriterFactory.createUserModelWriterWithTemplate("test-list.xlsx");
//		ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();

		List<MatterDto2> matterDtos = Lists.newArrayList();

		for (int i = 1; i <= 10; i++) {
			MatterDto2 dto = MatterDto2.builder().matterCode("120042981" + i).matterName("玻璃/钢化5*480*1528 \n test" + i).brand("apple" + i).num(100 + i).unit("PC" + i).build();
			matterDtos.add(dto);
		}

		writer.addModelList(matterDtos, "工作表1", true).excludes(new String[]{""});
		File file = new File("/Users/vachel/Downloads/test-list.xlsx");
		OutputStream outputStream = new FileOutputStream(file);

		Stopwatch stopwatch = Stopwatch.createStarted();
		writer.process(outputStream, "test-list", ExcelSuffixEnum.XLSX);
		long elapsed = stopwatch.elapsed(TimeUnit.MILLISECONDS);
		log.info(elapsed + "");

	}

	@Test
	public void testCache() throws Exception {
		WorkbookCachePool.init();
	}

	@Test
	public void shiftTest() throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("test");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("test");
		sheet.shiftRows(0,0,2);
		File file = new File("/Users/vachel/Downloads/test-list.xlsx");
		OutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
	}


}
