package com.github.liaoqn.excel.dto;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

/**
 * Excel sheet页列信息
 *
 * @author liaoqn
 * @date 2023/4/14
 */
@Getter
@Setter
public class SheetColumnData {

    /**
     * 哪个Sheet页
     */
    private int sheetNo;

    /**
     * 列信息
     */
    private List<ColumnData> columnDataList;

}
