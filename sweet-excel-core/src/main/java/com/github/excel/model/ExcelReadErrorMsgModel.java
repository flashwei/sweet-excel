package com.github.excel.model;

import com.github.excel.constant.ExcelConstant;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:06 下午
 * @Description: Excel 读取错误信息
 */
@Data
public class ExcelReadErrorMsgModel {
	private List<ExcelReadErrorMsgInfo> errorMsgInfoList = new ArrayList<>(ExcelConstant.TOW_INT);
	private Boolean existsError = false;

}
