package com.github.export;

import com.github.excel.write.ExcelWriterDataFormat;
import lombok.extern.slf4j.Slf4j;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:25 下午
 * @Description: 字段格式化
 */
@Slf4j
public class ExcelDefaultDataFormatTest implements ExcelWriterDataFormat {

	@Override
	public Object format(Object data, String pattern) {

		SimpleDateFormat sdf = new SimpleDateFormat(pattern);
		data = sdf.format((Date) data)+"时分秒";

		return data;
	}
}
