package com.github.excel.read;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:13 下午
 * @Description: 事件读取
 */
public interface ExcelEventRowReader<T> {

    /**
     * 业务逻辑实现方法
     * @param sheetIndex sheet序号 从0开始
     * @param curRow    当前行号
     * @param rValues   当前行数据，都已经转成string类型，无需关心excel细节，日期date 'yyyy-MM-dd HH:mm:ss'
     * @return
     */
    T getRows(int sheetIndex, int curRow, List<String> rValues);
}
