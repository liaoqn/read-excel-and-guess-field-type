package com.github.liaoqn.excel.dto;

import lombok.Getter;
import lombok.Setter;

/**
 * Excel sheet页
 *
 * @author liaoqn
 * @date 2023/4/14
 */
@Getter
@Setter
public class ExcelSheet {
    /**
     * sheet页编号，从0开始
     */
    private Integer sheetNo;
    /**
     * sheet name
     */
    private String sheetName;
}
