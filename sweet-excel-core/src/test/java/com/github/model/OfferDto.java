package com.github.model;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelExportListFillTypeEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Builder;
import lombok.Data;


/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 测试demo
 */
@Data
@ExcelExport(fillType = ExcelExportListFillTypeEnum.COVER,/*mergeContentRowNum = 1,mergeContentColNum = 1,mergeTitleColNum = 1,mergeTitleRowNum = 1, *//*fillStyle = ExcelExportFillStyleEnum.VERTICAL,titleStyleName = ExcelBasicStyle.STYLE_LIST_TITLE,contentStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT,*/theme = ExcelThemeEnum.ZEBRA)
@Builder
public class OfferDto extends ExcelBaseModel {
	@ExcelExportCell(titleName = "初始报价",colWidth = -1,rowHeight = -1)
	private Double initPrice;
	@ExcelExportCell(titleName = "最终报价",colWidth = -1,rowHeight = -1)
	private Double finalPrice;
	@ExcelExportCell(titleName = "净价",colWidth = -1,rowHeight = -1)
	private Double cleanPrice;
	@ExcelExportCell(titleName = "金额",colWidth = -1,rowHeight = -1)
	private Double price;
}
