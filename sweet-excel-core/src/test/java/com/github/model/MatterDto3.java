package com.github.model;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.enums.ExcelExportListFillTypeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelBasicStyle;
import lombok.Builder;
import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:29 下午
 * @Description: 测试demo
 */
@Data
@ExcelExport(rowIndex = 2,fillTitle = true,contentStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT,
		fillType = ExcelExportListFillTypeEnum.SHIFT, colIndex = 0,fillStyle = ExcelExportFillStyleEnum.VERTICAL,titleStyleName = ExcelBasicStyle.STYLE_LIST_TITLE)
@Builder
public class MatterDto3 extends ExcelBaseModel {

	@ExcelExportCell(titleName = "物料名称",colWidth = 200,rowHeight = -1)
	private String matterName;
	@ExcelExportCell(titleName = "物料编码",colWidth = 200,rowHeight = -1)
	private String matterCode;
	@ExcelExportCell(titleName = "品牌/材质/规格",colWidth = 200,rowHeight = -1)
	private String brand;
	@ExcelExportCell(titleName = "采购数量",colWidth = 200,rowHeight = -1)
	private String num;
	@ExcelExportCell(titleName = "单位",colWidth = 200,rowHeight = -1)
	private String unit;
}
