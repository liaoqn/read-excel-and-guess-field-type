package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.github.liaoqn.excel.enums.FieldType;

import java.math.BigDecimal;

/**
 * 转换为长整数
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class LongConverter extends AbstractNumberConverter<Long> {
    // 匹配货币 ￥20879。支持科学计数法（+），不支持小数
    // protected static final Pattern NUM_PATTERN = Pattern.compile("^[\\p{Sc}]?[\\s]*([-+]?\\d+(\\.?\\d*[Ee][+]?\\d+)?)$");

    public final static BigDecimal LONG_MAX = BigDecimal.valueOf(Long.MAX_VALUE);
    public final static BigDecimal LONG_MIN = BigDecimal.valueOf(Long.MIN_VALUE);

    @Override
    public FieldType destinationType() {
        return FieldType.LONG;
    }

    @Override
    protected Long doConvert(Object val, Cell cell) {
        BigDecimal decimalVal = convertToDecimal(val, cell);
        if (decimalVal == null) {
            return null;
        }

        if (decimalVal.compareTo(LONG_MAX) > 0) {
            return Long.MAX_VALUE;
        }
        if (decimalVal.compareTo(LONG_MIN) < 0) {
            return Long.MIN_VALUE;
        }

        return decimalVal.longValue();
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        BigDecimal decimalVal = convertToDecimal(val, cell);
        if (decimalVal == null) {
            return false;
        }

        return decimalVal.scale() <= 0 && decimalVal.compareTo(LONG_MAX) <= 0
                && decimalVal.compareTo(LONG_MIN) >= 0;
    }

}
