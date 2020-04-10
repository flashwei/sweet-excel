package com.github.excel.model;

import lombok.Data;
import org.apache.poi.ss.usermodel.Font;

import java.io.Serializable;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:06 下午
 * @Description: Excel 富文本
 */
@Data
public class ExcelRichTextModel {
	private Font font ;
	private int startIndex ;
	private int endIndex ;
}
