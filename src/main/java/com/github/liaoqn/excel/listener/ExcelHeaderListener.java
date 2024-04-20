package com.github.liaoqn.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.github.liaoqn.excel.dto.ColumnData;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * 读取excel表头
 *
 * @author liaoqn
 * @date 2023/4/19
 */
@Slf4j
public class ExcelHeaderListener implements ReadListener<Map<Integer, Object>> {

    private final Map<Integer, ColumnData> columnDataMap = new TreeMap<>();

    public ExcelHeaderListener() {
        super();
    }

    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        headMap.forEach((columnIndex, val) -> {
            ColumnData columnData = new ColumnData();
            columnData.setColumnIndex(columnIndex);
            columnData.setTitle(StringUtils.isEmpty(val.getStringValue()) ? null : String.valueOf(val.getStringValue()));
            columnDataMap.put(columnIndex, columnData);
        });
    }

    @Override
    public void invoke(Map<Integer, Object> data, AnalysisContext context) {

    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }

    public List<ColumnData> getColumnDataList() {
        return new ArrayList<>(columnDataMap.values());
    }
}
