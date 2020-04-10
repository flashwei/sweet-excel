package com.github.excel.model;

import lombok.Builder;
import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:05 下午
 * @Description: 表达式模型
 */
@Data
@Builder
public class ExcelExpressionModel {

    private static final long serialVersionUID = 5658780029129098912L;

    private String sheetName;

    private String nameSpace;

    private Integer rowIndex;

    private  Integer colIndex;

    private String  expression;

    private String  expressionContent;

    private String[] fieldName;

}
