package com.github.excel.model;

import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:09 下午
 * @Description: 读取图片model
 */
@Data
public class ReadPictureModel {
	private String sheetName;
	private Integer rowIndex;
	private Integer colIndex;
	private String point;
	private byte[] bytes;
	private String suffix;
}
