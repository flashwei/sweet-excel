package com.github.excel.read;

import com.google.common.collect.Lists;
import com.github.excel.exception.ExcelReadException;
import com.github.excel.read.handler.AbstractExecuteHandler;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.collections4.CollectionUtils;

import java.util.List;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:12 下午
 * @Description: 读取事件处理器
 */
public class ExcelEventReadExecutor<T> {
    /**
     * 是否批量处理
     */
    @Getter
    @Setter
    private Boolean batchExec;

    @Getter
    @Setter
    private Integer batchSize;

    private AbstractExecuteHandler<T> abstractExecuteHandler;

    private boolean everHaveProcessed = false;

    /**
     * 读取结果
     */
    private List<T> result;

    private List<String> errors;

    public ExcelEventReadExecutor(AbstractExecuteHandler<T> abstractExecuteHandler){
        init(abstractExecuteHandler.getBatchSize(),abstractExecuteHandler.getBatchExec());
        this.abstractExecuteHandler = abstractExecuteHandler;
    }

    public void init(Integer batchSize, Boolean batchExec){
        this.batchSize = batchSize;
        this.batchExec = batchExec;
        result = Lists.newArrayListWithCapacity(batchSize);
    }

    public void submit(T object){
        if (Objects.nonNull(object)) {
            result.add(object);
            if(result.size() >= batchSize && batchExec){
                this.batchExecuteAndClear();
            }
        }
    }

    public void flush(){
        this.batchExecuteAndClear();
    }

    /**
     * 执行处理器，并清空本次执行数据
     */
    private void batchExecuteAndClear(){
        try {
            if (!CollectionUtils.isEmpty(result)) {
                everHaveProcessed = true;
                abstractExecuteHandler.checks(result);
                abstractExecuteHandler.batchExecute(result);
                result.clear();
            }else {
                if(!everHaveProcessed){
                    throw new ExcelReadException("file.data.is.empty");
                }
            }
        } catch (Exception e) {
            result.clear();
            throw e;
        }
    }
}
