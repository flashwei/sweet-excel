package com.github.read.event;

import com.google.common.base.Strings;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.ExcelEventReader;
import com.github.excel.read.ExcelReaderFactory;
import com.github.excel.read.handler.AbstractExecuteHandler;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
public class ExcelEventImportTest {

    @Test
    public void test(){

        File  file = new File("/Users/terminus/Downloads/1535530298983.xlsx");

        ExcelEventReader<InquiryOrderItem> excelEventReader = ExcelReaderFactory.createExcelEventReader("1535530298983.xlsx");

        excelEventReader.setRowReader((sheetIndex,curRow,rowValue)->{
            InquiryOrderItem inquiryOrderItem = new InquiryOrderItem();
            if(curRow > 1){
                inquiryOrderItem = new InquiryOrderItem();
                inquiryOrderItem.setCompanyId(1L);

                if(Strings.isNullOrEmpty(rowValue.get(0))){
                    throw new ExcelReadException("item.name.is.empty");
                }
                inquiryOrderItem.setName(rowValue.get(0));

                inquiryOrderItem.setBrandName(rowValue.get(1));
                inquiryOrderItem.setSpec(rowValue.get(2));

                //价格
                if(Strings.isNullOrEmpty(rowValue.get(3))){
                    throw new ExcelReadException("item.quantity.is.empty");
                }
                //类型转换异常
                try{
                    inquiryOrderItem.setQuantity(Integer.parseInt(rowValue.get(3)));
                }catch (NumberFormatException e){
                    throw new ExcelReadException("item.quantity.is.illegal");
                }
                //单位
                if(Strings.isNullOrEmpty(rowValue.get(4))){

                    throw new ExcelReadException("item.unit.is.empty");
                }
                inquiryOrderItem.setUnit(rowValue.get(4));
            }
            return inquiryOrderItem;
        });

        excelEventReader.setExecuteHandler(new AbstractExecuteHandler<InquiryOrderItem>() {
            @Override
            public void batchExecute(List<InquiryOrderItem> result) {
                result.forEach(re -> {
                        System.out.println(re.toString());
                });
            }
        });

        try {
            excelEventReader.process(new FileInputStream(file));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }

}
