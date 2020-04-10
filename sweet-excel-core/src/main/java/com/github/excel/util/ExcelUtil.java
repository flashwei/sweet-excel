package com.github.excel.util;

import com.github.excel.boot.ExcelBootLoader;
import com.google.common.base.Throwables;
import com.github.excel.constant.ExcelConstant;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Objects;

@Slf4j
public class ExcelUtil {

    public static void setResponseHeader(HttpServletRequest request, HttpServletResponse response, String fileName, String suffix) throws UnsupportedEncodingException {
        if(StringUtil.notEmpty(suffix)) {
            fileName = fileName + suffix;
        }
        String userAgent = request.getHeader("USER-AGENT");
        String finalFileName;
        // 针对IE或者以IE为内核的浏览器：
        if (userAgent.contains("MSIE") || userAgent.contains("Trident")) {
            finalFileName = java.net.URLEncoder.encode(fileName, "UTF-8");
        } else {
            // 非IE浏览器的处理：
            finalFileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename=\"" + finalFileName + "\"");
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
    }

    /**
     * 设置下载header
     * @param response response
     * @param fileName 文件名称
     */
    public static void setResponseHeader(HttpServletResponse response, String fileName){
        if (StringUtil.isEmpty(fileName) || Objects.isNull(response)) {
            return ;
        }
        if (fileName.endsWith(ExcelConstant.XLSX_STR)) {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }else if (fileName.endsWith(ExcelConstant.XLSX_STR)) {
            response.setContentType("application/vnd.ms-excel");
        }
        try {
            fileName = URLEncoder.encode(fileName, "UTF8");
        } catch (UnsupportedEncodingException e) {
            log.error("Download import template error , cause:{}", Throwables.getStackTraceAsString(e));
        }
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
    }

    /**
     * 下载导入模板
     * @param response response
     * @param templateName 模板名称
     */
    public static void downloadImportTemplate(HttpServletResponse response, String templateName, String alias) {
        byte[] bytes = ExcelBootLoader.getExcelImportTemplateFileCacheMapValue(templateName);
        if (Objects.nonNull(bytes) && bytes.length > ExcelConstant.ZERO_SHORT) {
            if (StringUtil.isEmpty(alias)) {
                alias = templateName;
            }else{
                alias += templateName.substring(templateName.indexOf(ExcelConstant.DOT_CHAR));
            }
            setResponseHeader(response, alias);
            try {
                ServletOutputStream outputStream = response.getOutputStream();
                outputStream.write(bytes);
                outputStream.flush();
            } catch (IOException e) {
                log.error("Download import template error , cause:{}", Throwables.getStackTraceAsString(e));
            }
        }else{
            try {
                response.setHeader("Content-type", "text/html;charset=UTF-8");
                response.setCharacterEncoding("UTF-8");
                PrintWriter writer = response.getWriter();
                writer.print("<h1>下载模板失败，请确认模板【"+templateName+"】是否添加到导入模板文件目录！</h1>");
                writer.flush();
            } catch (IOException e) {
                log.error("Download import template error , cause:{}", Throwables.getStackTraceAsString(e));
            }
        }
    }

}
