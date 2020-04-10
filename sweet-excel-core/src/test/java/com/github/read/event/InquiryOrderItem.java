package com.github.read.event;

import lombok.Data;
import lombok.ToString;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.Date;
/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:30 下午
 * @Description: 简单导出
 */
@Data
@ToString
public class InquiryOrderItem implements Serializable {

    private static final long serialVersionUID = 2655052736801756779L;

    private Long id;

    private Long inquiryId;

    private Long companyId;

    private Long itemId;

    private Long skuId;

    private String erpCode;

    private String name;

    private String image;

    private Long brandId;

    private String brandName;

    private String spec;

    private Integer quantity;

    private String unit;

    private Long categoryId;

    private String fileName;

    private String filePath;

    private BigDecimal price;

    private Integer status;

    private Date createdAt;

    private Date updatedAt;

}
