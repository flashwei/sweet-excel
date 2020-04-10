package com.github.read.template;

import com.github.excel.read.AbstractImportTemplateExcluder;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
public class ImportTemplateExcluder extends AbstractImportTemplateExcluder {
	@Override
	public void addTemplateExclude() {
		this.addExclude(new SheetExclude("说明", "project-bids.xlsx")).addExclude(new RowExclude(1, "招标规划", "project-bids.xlsx"));
	}
}
