package com.github.excel.starter;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:19 下午
 * @Description: Sweet-Excel配置
 */
@Data
@ConfigurationProperties(prefix = "excel")
@Component
public class ExcelProperties {
    /**
     * 包路径
     */
    private List<String> dtoPackage;
    /**
     * 导出模板文件夹
     */
    private String exportTemplateDir;
    /**
     * 导入模板文件夹
     */
    private String importTemplateDir;
}
