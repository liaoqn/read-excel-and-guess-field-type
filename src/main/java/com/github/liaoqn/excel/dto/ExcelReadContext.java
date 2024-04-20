package com.github.liaoqn.excel.dto;

import com.github.liaoqn.excel.enums.FileType;
import lombok.Getter;
import lombok.Setter;

import java.io.InputStream;
import java.util.List;

/**
 * 读取excel入参
 *
 * @author liaoqn
 * @date 2023/4/18
 */
@Getter
@Setter
public class ExcelReadContext {
    /**
     * excel文件流
     */
    private InputStream inputStream;
    /**
     * excel类型
     */
    private FileType fileType;
    /**
     * sheet页编号，从0开始
     */
    private int sheetNo;
    /**
     * excel列信息
     */
    private List<ColumnData> columnDataList;
    /**
     * 从哪一行开始读取，从0开始计数。默认为1
     */
    private Integer offset;
    /**
     * 最多读取多少条。默认不限制
     */
    private Integer size;
}
