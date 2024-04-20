package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.github.liaoqn.excel.enums.FieldType;

import java.math.BigDecimal;

/**
 * 转换为整数
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class IntegerConverter extends AbstractNumberConverter<Integer> {

    public final static BigDecimal INT_MAX = BigDecimal.valueOf(Integer.MAX_VALUE);
    public final static BigDecimal INT_MIN = BigDecimal.valueOf(Integer.MIN_VALUE);

    @Override
    public FieldType destinationType() {
        return FieldType.INTEGER;
    }

    @Override
    protected Integer doConvert(Object val, Cell cell) {
        BigDecimal decimalVal = convertToDecimal(val, cell);
        if (decimalVal == null) {
            return null;
        }

        if (decimalVal.compareTo(INT_MAX) > 0) {
            return Integer.MAX_VALUE;
        }
        if (decimalVal.compareTo(INT_MIN) < 0) {
            return Integer.MIN_VALUE;
        }

        return decimalVal.intValue();
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        BigDecimal decimalVal = convertToDecimal(val, cell);
        if (decimalVal == null) {
            return false;
        }

        return decimalVal.scale() <= 0 && decimalVal.compareTo(INT_MAX) <= 0
                && decimalVal.compareTo(INT_MIN) >= 0;
    }

}
