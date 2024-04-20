package com.github.liaoqn.excel.dto;

import com.github.liaoqn.excel.enums.FieldType;
import lombok.Getter;
import lombok.Setter;

/**
 * excel列
 *
 * @author liaoqn
 * @date 2023/4/15
 */
@Getter
@Setter
public class ColumnData {
    /**
     * 列编号，从0开始
     */
    private int columnIndex;
    /**
     * 表头
     */
    private String title;
    /**
     * 字段类型
     */
    private FieldType fieldType;

    public ColumnData() {
    }

    public ColumnData(int columnIndex, String title, FieldType fieldType) {
        this.columnIndex = columnIndex;
        this.title = title;
        this.fieldType = fieldType;
    }
}
