package com.github.excel.util;

import com.google.common.collect.Maps;
import com.github.excel.constant.ExcelConstant;

import java.lang.reflect.Method;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:15 下午
 * @Description: 反射工具类
 */
public class ReflectCacherUtil {
	/**
	 * 缓存
	 */
	private Map<Class, Map<String, Method>> classCacheMap = Maps.newHashMap();

	/**
	 * 获取值并缓存class信息
	 * @param wrapObject
	 * @param attribute
	 * @return
	 * @throws Exception
	 */
	public Object getObjectThenCache(Object wrapObject, String attribute) throws Exception{
		if (Objects.isNull(wrapObject)) {
			return wrapObject;
		}
		char[] cs = attribute.toCharArray();
		cs[ExcelConstant.ZERO_SHORT] -= ExcelConstant.INT_32;
		String fieldName = String.valueOf(cs);
		String methodName = ExcelConstant.GET_STR + fieldName;
		Class<?> wrapObjectClass = wrapObject.getClass();
		cacheClassGetMethod(wrapObjectClass);
		Method method = classCacheMap.get(wrapObjectClass).get(methodName);
		if(Objects.nonNull(method)) {
			return method.invoke(wrapObject);
		}else{
			return null;
		}
	}

	/**
	 * 缓存class的method
	 * @param wrapObjectClass
	 */
	private void cacheClassGetMethod(Class<?> wrapObjectClass) {
		Map<String, Method> methodMap = classCacheMap.get(wrapObjectClass);
		if (Objects.isNull(methodMap)) {
			methodMap = Maps.newHashMap();
			for (Method method : wrapObjectClass.getDeclaredMethods()) {
				if(method.getName().startsWith(ExcelConstant.GET_STR)) {
					methodMap.put(method.getName(), method);
				}
			}
			classCacheMap.put(wrapObjectClass, methodMap);
		}
	}

}
