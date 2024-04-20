package com.github.liaoqn.excel.converter;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.github.liaoqn.excel.enums.FieldType;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.Optional;

/**
 * 转换为小数
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class DecimalConverter extends AbstractNumberConverter<BigDecimal> {

    @Override
    public FieldType destinationType() {
        return FieldType.DECIMAL;
    }

    @Override
    protected BigDecimal doConvert(Object val, Cell cell) {
        return convertToDecimal(val, cell);
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        String test = StringUtils.trimToNull(String.valueOf(val));
        if (StringUtils.isEmpty(test)) {
            return false;
        }

        if (NumberUtils.isParsable(test)) {
            return true;
        }

        return Optional.ofNullable(cell)
                .filter(ReadCellData.class::isInstance)
                .map(ReadCellData.class::cast)
                .filter(c -> CellDataTypeEnum.NUMBER.equals(c.getType()))
                .isPresent();
    }
}
