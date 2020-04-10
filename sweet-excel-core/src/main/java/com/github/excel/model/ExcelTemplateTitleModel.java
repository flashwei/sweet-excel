package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

import java.io.Serializable;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:08 下午
 * @Description: 模板标题模型
 */
@Data
@Builder
@AllArgsConstructor
public class ExcelTemplateTitleModel {

    private String sheetName;

    private Integer rowIndex;

    private  Integer colIndex;

    private String title;

}
