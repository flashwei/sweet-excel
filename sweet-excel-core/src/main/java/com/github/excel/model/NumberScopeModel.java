package com.github.excel.model;

import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:09 下午
 * @Description: 导出验证-数字范围
 */
@Data
public final class NumberScopeModel extends CommentModel{
	private String start;
	private String end;
}
