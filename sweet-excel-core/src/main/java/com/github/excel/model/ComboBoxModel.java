package com.github.excel.model;

import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:02 下午
 * @Description: 导出验证-下拉框model
 */
@Data
public final class ComboBoxModel extends CommentModel{
	private String[] options;
}
