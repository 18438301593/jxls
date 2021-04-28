package com.demo.entity;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 封装实体类
 */
@Data
@AllArgsConstructor
public class Reports {
    //序号
    private Integer num;
    //项目名称
    private String name;
    //责任单位
    private String company;
    //项目来源
    private String sources;
    //总投资
    private Double count;
}
