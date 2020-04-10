package com.github.excel.model;

import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:02 下午
 * @Description: 批注model
 */
@Data
public class CommentModel {
	private Object value;
	private String commentFontName;
	private String commentText;
}
