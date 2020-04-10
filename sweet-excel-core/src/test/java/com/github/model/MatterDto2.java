package com.github.model;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.enums.ExcelExportFillStyleEnum;
import com.github.excel.enums.ExcelExportListFillTypeEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelBasicStyle;
import lombok.Builder;
import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:28 下午
 * @Description: 测试demo
 */
@Data
@ExcelExport(rowIndex = 2,colIndex = 0,fillTitle = false,fillType = ExcelExportListFillTypeEnum.SHIFT, fillStyle = ExcelExportFillStyleEnum.VERTICAL,titleStyleName = ExcelBasicStyle.STYLE_ZEBRA_TITLE_ROW,theme = ExcelThemeEnum.ZEBRA)
@Builder
public class MatterDto2 extends ExcelBaseModel {

	@ExcelExportCell(titleName = "物料名称")
	private String matterName;
	@ExcelExportCell(titleName = "物料编码")
	private String matterCode;
	@ExcelExportCell(titleName = "品牌/材质/规格")
	private String brand;
	@ExcelExportCell(titleName = "采购数量")
	private Integer num;
	@ExcelExportCell(titleName = "单位")
	private String unit;
}
