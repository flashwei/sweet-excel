package com.github.excel.boot;

import com.github.excel.annotation.ExcelExport;
import com.github.excel.annotation.ExcelExportCell;
import com.github.excel.model.*;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.github.excel.annotation.ExcelImport;
import com.github.excel.annotation.ExcelImportProperty;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.AbstractImportTemplateExcluder;
import com.github.excel.util.PackageUtil;
import com.github.excel.util.StringUtil;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.Clock;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 6:57 下午
 * @Description: Excel 引导加载器
 */
@Slf4j
public class ExcelBootLoader {
    /**
     * 导出缓存
     */
    private static final Map<Class, ExcelCacheModel> EXCEL_CACHE_MAP = Maps.newConcurrentMap();
    /**
     * 导入缓存
     */
    private static final Map<Class, ExcelCacheImportModel> EXCEL_CACHE_IMPORT_MAP = Maps.newConcurrentMap();

    /**
     * 模板表达式缓存
     */
    private static final Map<String, Map<String, List<ExcelExpressionModel>>> EXCEL_TEMPLATE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 模板标题缓存
     */
    private static final Map<String, Map<Integer, List<ExcelTemplateTitleModel>>> EXCEL_TEMPLATE_TITLE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 导出模板缓存
     */
    private static final Map<String, ExcelTemplateCacheModel> EXCEL_TEMPLATE_FILE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 导入模板缓存
     */
    private static final Map<String, byte[]> EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 导入模板缓存 key->文件名, value-> (key->sheet名,value->模板内容)
     */
    private static final Map<String, Map<String,List<ExcelImportTemplateCacheModel>>> EXCEL_IMPORT_TEMPLATE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 模板文件夹地址
     */
    private static String TEMPLATE_DIR_PATH ;

    /**
     * 获取导出模板文件夹名称
     * @return
     */
    public static String getTemplateDirPath(){
        return TEMPLATE_DIR_PATH;
    }

    /**
     * 获取导出缓存
     * @param key 模板名称
     * @return
     */
    public static ExcelCacheModel getExcelCacheMapValue(Class key){
        return EXCEL_CACHE_MAP.get(key);
    }

    /**
     * 获取导出模板表达式缓存
     * @param key 模板名称
     * @return
     */
    public static Map<String, List<ExcelExpressionModel>> getExcelTemplateCacheMapValue(String key){
        return EXCEL_TEMPLATE_CACHE_MAP.get(key);
    }

    /**
     * 获取导入模板缓存
     * @param key 模型class
     * @return
     */
    public static ExcelCacheImportModel getExcelCacheImportMapValue(Class key){
        return EXCEL_CACHE_IMPORT_MAP.get(key);
    }

    /**
     * 获取导出标题缓存
     * @param key 模板名称
     * @return
     */
    public static Map<Integer, List<ExcelTemplateTitleModel>> getExcelTemplateTitleCacheMapValue(String key){
        return EXCEL_TEMPLATE_TITLE_CACHE_MAP.get(key);
    }

    /**
     * 获取导出模板文件缓存(excel)
     * @param key
     * @return
     */
    public static ExcelTemplateCacheModel getExcelTemplateFileCacheMapValue(String key){
        return EXCEL_TEMPLATE_FILE_CACHE_MAP.get(key);
    }

    /**
     * 获取导入模板文件流
     * @param key 导入模板文件名
     * @return
     */
    public static byte[] getExcelImportTemplateFileCacheMapValue(String key){
        return EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP.get(key);
    }

    /**
     * 获取导入模板缓存
     * @param key 模板文件名称
     * @return
     */
    public static Map<String,List<ExcelImportTemplateCacheModel>> getExcelImportTemplateCacheMapValue(String key){
        return EXCEL_IMPORT_TEMPLATE_CACHE_MAP.get(key);
    }


    /**
     * 加载并缓存excel模型信息
     *
     * @param packagePathArray
     */
    @SuppressWarnings("unchecked")
    public static void loadModel(String ... packagePathArray) {
        if (Objects.isNull(packagePathArray) || packagePathArray.length == ExcelConstant.ZERO_SHORT) {
            throw new IllegalArgumentException("packagePath can't be null");
        }
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("Sweet Excel loading model");
        List<Class<?>> classList = new ArrayList<>();
        for(String packagePath:packagePathArray) {
            List<Class<?>> clasList = PackageUtil.getClass(packagePath, true);
            classList.addAll(clasList);
        }
        for (Class cla : classList) {
            ExcelExport excelExport = (ExcelExport) cla.getDeclaredAnnotation(ExcelExport.class);
            ExcelImport excelImport = (ExcelImport) cla.getDeclaredAnnotation(ExcelImport.class);
            if ((Objects.isNull(excelExport) && Objects.isNull(excelImport)) || !ExcelBaseModel.class.isAssignableFrom(cla)) {
                continue;
            }
            if(Objects.nonNull(excelExport)) {
                loadExportModel(cla, excelExport);
            }else if (Objects.nonNull(excelImport)){
                loadImportModel(cla, excelImport);
            }
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("Sweet Excel load success by {} millisecond", loadMillis);
    }

    /**
     * @Author: Vachel Wang
     * @Date: 2020/4/2 6:57 下午
     * @Description: 加载导出模型
     */
    private static void loadExportModel(Class cla, ExcelExport excelExport) {
        ExcelCacheModel cacheModel = new ExcelCacheModel();
        cacheModel.setExcelExport(excelExport);
        List<ExcelCacheModel.ExcelCacheFieldModel> fieldModelList = Lists.newArrayList();

        setIncrementSequenceNoField(excelExport,fieldModelList);

        Map<String,ExcelCacheModel.ExcelCacheFieldModel> fieldModelMap = Maps.newHashMap();
        for (Field excelField : cla.getDeclaredFields()) {
            ExcelExportCell exportCell = excelField.getDeclaredAnnotation(ExcelExportCell.class);
            if (Objects.isNull(exportCell) || exportCell.disable()) {
                continue;
            }
            ExcelCacheModel.ExcelCacheFieldModel fieldModel = new ExcelCacheModel.ExcelCacheFieldModel();

            short titleRowHeight = exportCell.rowHeight();
            short contentRowHeight = exportCell.rowHeight();
            short colWidth = exportCell.colWidth();
            if (excelExport.theme() != ExcelThemeEnum.NONE) {
                if (ExcelConstant.MINUS_TWO_SHORT == contentRowHeight) {
                    contentRowHeight = excelExport.theme().getContentRowHeight();
                    titleRowHeight = excelExport.theme().getTitleRowHeight();
                }
                if (ExcelConstant.MINUS_TWO_SHORT == colWidth) {
                    colWidth = excelExport.theme().getColWidth();
                }
            }
            //当属性是Map类型的时候  设置个标记
            if(excelField.getType().equals(Map.class)){
                fieldModel.setMap(true);
            }else {
                fieldModel.setMap(false);
            }
            char[] cs = excelField.getName().toCharArray();
            cs[ExcelConstant.ZERO_SHORT] -= ExcelConstant.INT_32;
            String fieldName = String.valueOf(cs);
            try {
                Method getMethod = cla.getMethod(ExcelConstant.GET_STR + fieldName);
                boolean isDate = false;
                if (Date.class.isAssignableFrom(getMethod.getReturnType()) || Calendar.class.isAssignableFrom(getMethod.getReturnType())) {
                    isDate = true;
                }
                String titleStyleName = StringUtil.notEmpty(exportCell.titleStyleName()) ? exportCell.titleStyleName() : StringUtil.notEmpty(excelExport.titleStyleName()) ? excelExport.titleStyleName() : excelExport.theme().getTitleRowStyleName();
                String contentStyleName = StringUtil.notEmpty(exportCell.contentStyleName()) ? exportCell.contentStyleName() : StringUtil.notEmpty(excelExport.contentStyleName()) ? excelExport.contentStyleName() : isDate ? excelExport.theme().getOddRowStyleDateName():excelExport.theme().getOddRowStyleName();
                String evenRowStyleName = StringUtil.notEmpty(exportCell.contentStyleName()) ? exportCell.contentStyleName() : StringUtil.notEmpty(excelExport.contentStyleName()) ? excelExport.contentStyleName() : isDate ? excelExport.theme().getEvenRowStyleDateName():excelExport.theme().getEvenRowStyleName();
                fieldModel.setExportCell(exportCell);
                fieldModel.setGetMethod(getMethod);
                fieldModel.setFieldName(excelField.getName());
                fieldModel.setTitleStyleName(titleStyleName);
                fieldModel.setContentStyleName(contentStyleName);
                fieldModel.setEvenRowStyleName(evenRowStyleName);
                fieldModel.setTitleRowHeight(titleRowHeight);
                fieldModel.setContentRowHeight(contentRowHeight);
                fieldModel.setColWidth(colWidth);
                fieldModelList.add(fieldModel);
                fieldModelMap.put(exportCell.titleName(), fieldModel);
            } catch (NoSuchMethodException e) {
                log.error(Throwables.getStackTraceAsString(e));
                throw new ExcelReadException("load export model failed ,cause:" + e.getMessage());
            }
        }
        cacheModel.setFieldModelList(fieldModelList);
        cacheModel.setFieldModelMap(fieldModelMap);
        EXCEL_CACHE_MAP.put(cla, cacheModel);
    }

    private static void setIncrementSequenceNoField(ExcelExport excelExport , List<ExcelCacheModel.ExcelCacheFieldModel> fieldModelList){
        if (excelExport.incrementSequenceNo()) {

            String titleStyleName =  StringUtil.notEmpty(excelExport.titleStyleName()) ? excelExport.titleStyleName() : excelExport.theme().getTitleRowStyleName();
            String contentStyleName = StringUtil.notEmpty(excelExport.contentStyleName()) ? excelExport.contentStyleName() : excelExport.theme().getOddRowStyleName();
            String evenRowStyleName = StringUtil.notEmpty(excelExport.contentStyleName()) ? excelExport.contentStyleName() : excelExport.theme().getEvenRowStyleName();

            ExcelCacheModel.ExcelCacheFieldModel fieldModel = new ExcelCacheModel.ExcelCacheFieldModel();
            fieldModel.setMap(false);
            fieldModel.setFieldName(ExcelConstant.INCREMENT_SEQUENCE_NO_FIELD_NAME);
            fieldModel.setTitleStyleName(titleStyleName);
            fieldModel.setContentStyleName(contentStyleName);
            fieldModel.setEvenRowStyleName(evenRowStyleName);
            fieldModel.setColWidth(ExcelConstant.SHORT_100);
            fieldModel.setTitleRowHeight(ExcelConstant.MINUS_TWO_SHORT);
            fieldModel.setContentRowHeight(ExcelConstant.MINUS_TWO_SHORT);
            fieldModelList.add(fieldModel);
        }
    }

    /**
     * 加载导入模型
     * @param cla
     * @param excelImport
     */
    private static void loadImportModel(Class cla, ExcelImport excelImport) {
        ExcelCacheImportModel cacheModel = new ExcelCacheImportModel();
        cacheModel.setExcelImport(excelImport);
        Map<String,ExcelCacheImportModel.ExcelCacheImportFieldModel> fieldModelMap = Maps.newHashMap();
        for (Field excelField : cla.getDeclaredFields()) {
            ExcelImportProperty importProperty = excelField.getDeclaredAnnotation(ExcelImportProperty.class);
            if (Objects.isNull(importProperty) || importProperty.disable()) {
                continue;
            }
            ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = new ExcelCacheImportModel.ExcelCacheImportFieldModel();
            char[] cs = excelField.getName().toCharArray();
            cs[ExcelConstant.ZERO_SHORT] -= ExcelConstant.INT_32;
            String fieldName = String.valueOf(cs);
            try {
                Method setMethod = cla.getMethod(ExcelConstant.SET_STR + fieldName, excelField.getType());
                fieldModel.setImportProperty(importProperty);
                fieldModel.setSetMethod(setMethod);
                String modelMapKey = excelImport.enableSeparator() ? importProperty.titleName() + importProperty.separator() : importProperty.titleName();
                fieldModelMap.put(modelMapKey, fieldModel);
            } catch (NoSuchMethodException e) {
                log.error(Throwables.getStackTraceAsString(e));
                throw new ExcelReadException("load import model failed ,cause:" + e.getMessage());
            }
        }
        cacheModel.setFieldModelMap(fieldModelMap);
        EXCEL_CACHE_IMPORT_MAP.put(cla, cacheModel);
    }


    /**
     *  解析模板并缓存
     */
    public static void loadExcelTemplate(String templatePath) {
        TEMPLATE_DIR_PATH = templatePath+=ExcelConstant.FILE_SEPARATOR;
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("Sweet Excel loading export template");
        try {
            //读所有的文件
            File directory = new File(templatePath);
            if (directory.isDirectory()) {
                File[] files = Optional.ofNullable(directory.listFiles()).orElse(new File[ExcelConstant.ZERO_SHORT]);
                List<File> fileList = getFileList(files);
                ExcelTemplateLoader.loadExportTemplate(fileList,EXCEL_TEMPLATE_CACHE_MAP,EXCEL_TEMPLATE_TITLE_CACHE_MAP,EXCEL_TEMPLATE_FILE_CACHE_MAP);
            }else {
                log.error("directory.is.empty");
                throw new ExcelReadException("directory.is.empty");
            }
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReadException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("Sweet Excel load template success by {} millisecond", loadMillis);
    }

    /**
     *  导入模板解析并缓存
     */
    public static void loadImportExcelTemplate(String templatePath, AbstractImportTemplateExcluder excluder) {
        TEMPLATE_DIR_PATH = templatePath+=ExcelConstant.FILE_SEPARATOR;
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("Sweet Excel loading import template");
        try {
            //读所有的文件
            File directory = new File(templatePath);
            if (directory.isDirectory()) {
                File[] files = Optional.ofNullable(directory.listFiles()).orElse(new File[ExcelConstant.ZERO_SHORT]);
                List<File> fileList = getFileList(files);
                ExcelTemplateLoader.loadImportTemplate(fileList,EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP,EXCEL_IMPORT_TEMPLATE_CACHE_MAP,excluder);
            }else {
                log.error("directory.is.empty");
                throw new ExcelReadException("directory.is.empty");
            }
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReadException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("Sweet Excel load template success by {} millisecond", loadMillis);
    }

    /**
     *  解析模板并缓存
     */
    public static void loadExcelTemplate(List<Map<String, Object>> fileList) {
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("Sweet Excel loading export template");
        try {
            //读所有的文件
            ExcelTemplateLoader.loadBootExportTemplate(fileList,EXCEL_TEMPLATE_CACHE_MAP,EXCEL_TEMPLATE_TITLE_CACHE_MAP,EXCEL_TEMPLATE_FILE_CACHE_MAP);
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReadException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("Sweet Excel load template success by {} millisecond", loadMillis);
    }

    /**
     *  导入模板解析并缓存
     */
    public static void loadImportExcelTemplate(List<Map<String, Object>> fileList, AbstractImportTemplateExcluder excluder) {
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("Sweet Excel loading import template");
        try {
            //读所有的文件
            ExcelTemplateLoader.loadBootImportTemplate(fileList,EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP,EXCEL_IMPORT_TEMPLATE_CACHE_MAP,excluder);
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReadException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("Sweet Excel load template success by {} millisecond", loadMillis);
    }

    /**
     * @Author: Vachel Wang
     * @Date: 2020/4/2 6:58 下午
     * @Description: 缓存workbook
     */
    public static void cacheWorkbook() {
        WorkbookCachePool.init();
    }

    /**
     * 获取模板文件list
     * @param files
     * @return
     */
    private static List<File> getFileList(File[] files) {
        List<File> fileList = new ArrayList<>();
        if (Objects.isNull(files) || files.length == ExcelConstant.ZERO_SHORT) {
            return fileList;
        }
        for (File file : files) {
            String templateName = file.getName();
            if (templateName.endsWith(ExcelConstant.XLSX_STR) || templateName.endsWith(ExcelConstant.XLS_STR)) {
                fileList.add(file);
            }
        }
        return fileList;
    }

}
