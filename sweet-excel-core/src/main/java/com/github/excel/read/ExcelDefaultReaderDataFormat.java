package com.github.excel.read;

import com.github.excel.model.ReadPictureModel;
import com.github.excel.util.StringUtil;
import com.github.excel.constant.ExcelConstant;
import lombok.extern.slf4j.Slf4j;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:12 下午
 * @Description: 读取字段格式化
 */
@Slf4j
public class ExcelDefaultReaderDataFormat implements ExcelReaderDataFormat {

	@Override
	public Object format(Object data, String pattern, Class<?> targetCla) throws ParseException {
		Object result = data;
		if (Objects.isNull(result) || result instanceof ReadPictureModel) {
			return result;
		}
		if (targetCla == Date.class) {
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_DATE_FORMAT;
				}
				SimpleDateFormat sdf = new SimpleDateFormat(pattern);
				result = sdf.parse((String) data);
			}
		} else if (targetCla == Calendar.class) {
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_DATE_FORMAT;
				}
				SimpleDateFormat sdf = new SimpleDateFormat(pattern);
				Date date = sdf.parse((String) data);
				Calendar calendar = Calendar.getInstance();
				calendar.setTime(date);
				result = calendar;
			}
		} else if (Number.class.isAssignableFrom(targetCla)) {
			DecimalFormat df = StringUtil.isEmpty(pattern) ? new DecimalFormat() : new DecimalFormat(pattern);
			Number number = df.parse(data.toString());
			if (targetCla == Integer.class) {
				result = number.intValue();
			} else if (targetCla == Long.class) {
				result = number.longValue();
			} else if (targetCla == BigInteger.class) {
				result = new BigInteger(String.valueOf(number.longValue()));
			} else if (targetCla == Short.class) {
				result = number.shortValue();
			} else if (targetCla == Byte.class) {
				result = number.byteValue();
			} else if (targetCla == Float.class) {
				result = number.floatValue();
			} else if (targetCla == Double.class) {
				result = number.doubleValue();
			} else if (targetCla == BigDecimal.class) {
				result = new BigDecimal(String.valueOf(number.doubleValue()));
			}
		} else if (targetCla == Boolean.class) {
			if (data instanceof String) {
				result = Boolean.valueOf((String) data);
			}
		}
		return result;
	}
}
