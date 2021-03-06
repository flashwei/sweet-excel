package com.github.model;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelBasicStyle;
import lombok.Builder;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:28 下午
 * @Description: 测试demo
 */
@Data
@ExcelExport(nameSpace = "matter",incrementSequenceNo = true,incrementSequenceTitle = "编号啊",rowIndex = 0, colIndex = 0,fillStyle = ExcelExportFillStyleEnum.VERTICAL,titleStyleName = ExcelBasicStyle.STYLE_LIST_TITLE,contentStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT)
@Builder
public class MatterDto extends ExcelBaseModel {
	@ExcelExportCell(titleName = "物料编码",colWidth = 200,rowHeight = -1)
	private String matterCode;
	@ExcelExportCell(titleName = "物料名称", contentStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT,colWidth = 200,rowHeight = -1)
	private String matterName;
	@ExcelExportCell(titleName = "规格/品牌",colWidth = 200,rowHeight = -1)
	private String brand;
	@ExcelExportCell(titleName = "配置/材质/重量",colWidth = 200,rowHeight = -1)
	private String weight;
	@ExcelExportCell(titleName = "单位",colWidth = 200,rowHeight = -1)
	private String unit;
	@ExcelExportCell(titleName = "供应商",colWidth = 200,rowHeight = -1)
	private String supplier;
	@ExcelExportCell(titleName = "上次采购价格",colWidth = 200,rowHeight = -1)
	private Double prevPurchasePrice;
	@ExcelExportCell(titleName = "初始报价",colWidth = 200,rowHeight = -1)
	private Double initialPrice;
	@ExcelExportCell(titleName = "最终成交价",colWidth = 200,rowHeight = -1)
	private Double finalPrice;
	@ExcelExportCell(titleName = "税率",colWidth = 200,rowHeight = -1)
	private Double rate;
	@ExcelExportCell(titleName = "交货期/采购提前期",colWidth = 200,rowHeight = -1)
	private Integer deliveryDate;
	@ExcelExportCell(titleName = "运费承担",colWidth = 200,rowHeight = -1)
	private String freight;
	@ExcelExportCell(titleName = "付款方式",colWidth = 200,rowHeight = -1)
	private String payType;
	@ExcelExportCell(titleName = "联系方式",colWidth = 200,rowHeight = -1)
	private String concatMobile;
	@ExcelExportCell(titleName = "联系人",colWidth = 200,rowHeight = -1)
	private String concatName;
	@ExcelExportCell(titleName = "建议采购",colWidth = 200,rowHeight = -1)
	private String suggest;
	@ExcelExportCell(titleName = "备注",colWidth = 200,rowHeight = -1)
	private String remark;
	@ExcelExportCell(titleName = "询价时间",colWidth = 200,rowHeight = -1,formatPattern = "yyyy-MM-dd")
	private Date queryTime;

}
