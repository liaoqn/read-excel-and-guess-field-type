package com.github.liaoqn.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.github.liaoqn.excel.converter.ConverterFactory;
import com.github.liaoqn.excel.dto.ColumnData;
import com.github.liaoqn.excel.enums.FieldType;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collectors;

/**
 * 读取excel表头，并推测所在列的类型
 *
 * @author liaoqn
 * @date 2023/4/19
 */
@Slf4j
public class ExcelMetadataListener extends AnalysisEventListener<Map<Integer, Object>> {

    private final Map<Integer, Set<FieldType>> types = new HashMap<>();

    private final Map<Integer, ColumnData> columnDataMap = new TreeMap<>();

    /**
     * 本次读取的总条数
     */
    private final int size;
    private final Integer sheetNo;
    private int count;

    public ExcelMetadataListener(Integer size) {
        this.size = (size == null || size <= 0) ? 10 : size;
        this.count = 0;
        this.sheetNo = null;
    }

    public ExcelMetadataListener(Integer size, int sheetNo) {
        this.size = (size == null || size <= 0) ? 10 : size;
        this.count = 0;
        this.sheetNo = sheetNo;
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
        count = count + 1;

        data.forEach((columnIndex, val) -> {
            Map<Integer, Cell> cellMap = context.readRowHolder().getCellMap();
            Cell cell = cellMap.get(columnIndex);

            FieldType fieldType = ConverterFactory.guessType(val, cell);
            Set<FieldType> fieldTypes = types.computeIfAbsent(columnIndex, k -> new HashSet<>());
            fieldTypes.add(fieldType);
        });
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }

    @Override
    public boolean hasNext(AnalysisContext context) {
        // 超过size，不再读取
        return count < size;
    }

    public List<ColumnData> getColumnDataList() {
        if (MapUtils.isEmpty(types)) {
            return new ArrayList<>(columnDataMap.values());
        }

        return columnDataMap.values().stream().peek(columnData -> {
            Set<FieldType> fieldTypes = types.get(columnData.getColumnIndex());
            FieldType fieldType = destineType(fieldTypes);
            columnData.setFieldType(fieldType);
        }).collect(Collectors.toList());
    }

    public Integer getSheetNo() {
        return sheetNo;
    }

    private FieldType destineType(Set<FieldType> types) {
        if (CollectionUtils.isEmpty(types)) {
            return FieldType.STRING;
        }
        if (types.size() == 1) {
            return types.iterator().next();
        }
        if (types.stream().allMatch(e -> FieldType.LONG == e || FieldType.INTEGER == e)) {
            return FieldType.LONG;
        }
        if (types.stream().allMatch(e -> FieldType.LONG == e || FieldType.INTEGER == e || FieldType.DECIMAL == e)) {
            return FieldType.DECIMAL;
        }
        if (types.stream().anyMatch(e -> FieldType.LONG_TEXT == e)) {
            return FieldType.LONG_TEXT;
        }

        return FieldType.STRING;
    }
}
