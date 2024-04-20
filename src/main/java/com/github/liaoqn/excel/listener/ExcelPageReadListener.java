package com.github.liaoqn.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.util.ListUtils;
import com.github.liaoqn.excel.converter.ConverterFactory;
import com.github.liaoqn.excel.dto.ColumnData;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * 分页读取
 *
 * @author liaoqn
 * @date 2023/4/25
 */
@Slf4j
public class ExcelPageReadListener extends AnalysisEventListener<Map<Integer, Object>> {

    private final List<ColumnData> columnDataList;
    /**
     * 分页大小
     */
    private final int pageSize;
    /**
     * 读取到pageSize条数据后，触发回调
     */
    private final Consumer<List<ArrayList<Object>>> pageConsumer;
    private List<ArrayList<Object>> currPageCache;

    public ExcelPageReadListener(List<ColumnData> columnDataList, int pageSize, Consumer<List<ArrayList<Object>>> pageConsumer) {
        this.columnDataList = columnDataList;
        this.pageSize = pageSize;
        this.pageConsumer = pageConsumer;

        this.resetCurrPageCache();
    }

    @Override
    public void invoke(Map<Integer, Object> data, AnalysisContext context) {
        ArrayList<Object> row = new ArrayList<>(this.columnDataList.size());
        for (int i = 0; i < this.columnDataList.size(); i++) {
            ColumnData columnData = this.columnDataList.get(i);
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

        this.currPageCache.add(row);

        if (this.currPageCache.size() >= this.pageSize) {
            // 达到分页大小，触发回调
            this.pageConsumer.accept(this.currPageCache);

            // 回调执行完，清空页缓存数据
            this.resetCurrPageCache();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        if (CollectionUtils.isNotEmpty(this.currPageCache)) {
            this.pageConsumer.accept(this.currPageCache);
        }
    }

    /**
     * 清空页缓存数据
     */
    private void resetCurrPageCache() {
        this.currPageCache = ListUtils.newArrayListWithExpectedSize(this.pageSize);
    }
}
