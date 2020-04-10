package com.github.read.template;

import com.sun.istack.internal.NotNull;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import javax.annotation.Nullable;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
@Slf4j
public class ExcelReaderTest {


    @Test
    public void Test2() throws IOException {
        //ClassPath classPath = new ClassPath();

//        String path = ExcelLoadModelBoot.class.getClassLoader().getResource("excel-template").getPath();
        //URL url = boot.getClass().getClassLoader().getResource("/excel-template/模板测试.xlsx");

//        ExcelLoadModelBoot.loadExcelTemplate( path);

        /*Map<String, List<ExcelExpressionModel>> listMap = ExcelLoadModelBoot.EXCEL_TEMPLATE_EXPRESSION_CACHE_MAP.get("零星1模板-1534225773887.xls");

        listMap.forEach((k,v)->{
            v.forEach(e->{
                System.out.println(e.getExpression());
            });
        });*/

        //File f = new File("/Users/terminus/Ideaworkspace/Sweet-Excel/target/test-classes/excel-template");
        //f.isDirectory();
        String aaaa = test(null);
    }

    public @Nullable String test(@NotNull String str){
        System.out.println(str);
        return str ;
    }

    @Test
    public void Test3() throws IOException {
        String regex = "\\$\\{[A-Z,a-z,0-9]+((\\.{1})([A-Z,a-z,0-9]+))+\\}";
        String conent = "姓名：${aa1.bb2.cc3}，年龄：${aa1.bb2.cc4}，年龄：${aa1.bb2.2555}";
//        String conent = "111${aa1.bb2.cc3}111";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(conent);
        int i = matcher.groupCount();
        while (matcher.find()) {
            System.out.println(matcher.group(0));
        }
        System.out.println("${aa.bb}".matches(regex));
    }

}
