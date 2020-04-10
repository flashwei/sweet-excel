package com.github.excel.write;

import lombok.extern.slf4j.Slf4j;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:17 下午
 * @Description: 字段格式化
 */
@Slf4j
public class ExcelDefaultWriterDataFormat implements ExcelWriterDataFormat {

	@Override
	public Object format(Object data, String pattern) {
		if (data instanceof Date) {
			SimpleDateFormat sdf = new SimpleDateFormat(pattern);
			data = sdf.format((Date) data);
		} else if (data instanceof Calendar) {
			Calendar calendar = (Calendar) data;
			SimpleDateFormat sdf = new SimpleDateFormat(pattern);
			data = sdf.format(calendar.getTime());
		} else if (data instanceof Number) {
			DecimalFormat df = new DecimalFormat(pattern);
			Number number = (Number)data;
			data = df.format(number.doubleValue());
		}

		return data;
	}
}
