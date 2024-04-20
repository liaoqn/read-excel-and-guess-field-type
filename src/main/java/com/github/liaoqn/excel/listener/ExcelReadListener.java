package com.github.liaoqn.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Cell;
import com.github.liaoqn.excel.converter.ConverterFactory;
import com.github.liaoqn.excel.dto.ColumnData;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelReadListener extends AnalysisEventListener<Map<Integer, Object>> {

    private final List<ArrayList<Object>> dataList = new ArrayList<>();

    private final List<ColumnData> columnDataList;

    /**
     * 本次读取的总条数
     */
    private final Integer size;

    public ExcelReadListener(List<ColumnData> columnDataList, Integer size) {
        this.columnDataList = columnDataList;
        this.size = size;
    }

    @Override
    public void invoke(Map<Integer, Object> data, AnalysisContext context) {
        ArrayList<Object> row = new ArrayList<>(columnDataList.size());
        for (int i = 0; i < columnDataList.size(); i++) {
            ColumnData columnData = columnDataList.get(i);
            int columnIndex = columnData.getColumnIndex();
            Object val = data.get(columnIndex);
            if (val == null) {
                row.add(null);
                continue;
            }

            Map<Integer, Cell> cellMap = context.readRowHolder().getCellMap();
            Cell cell = cellMap.get(columnIndex);

            Object destinationTypeValue = ConverterFactory.getConvert(columnData.getFieldType()).convert(val, cell);
            row.add(destinationTypeValue);
        }

        dataList.add(row);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }

    @Override
    public boolean hasNext(AnalysisContext context) {
        // 超过size，不再读取
        return size == null || dataList.size() < size;
    }

    public List<ArrayList<Object>> getDataList() {
        return dataList;
    }
}
