package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.github.liaoqn.excel.enums.FieldType;

import java.math.BigDecimal;
import java.util.Optional;

/**
 * @author liaoqn
 * @date 2023/4/18
 */
public class ConverterFactory {
    private static final StringConverter STRING_CONVERTER = new StringConverter();
    private static final IntegerConverter INTEGER_CONVERTER = new IntegerConverter();
    private static final LongConverter LONG_CONVERTER = new LongConverter();
    private static final DateConverter DATE_CONVERTER = new DateConverter();
    private static final BoolConverter BOOL_CONVERTER = new BoolConverter();
    private static final DecimalConverter DECIMAL_CONVERTER = new DecimalConverter();
    private static final LongTextConverter LONG_TEXT_CONVERTER = new LongTextConverter();
    private static final NullConverter NULL_CONVERTER = new NullConverter();

    public static Converter<?> getConvert(FieldType fieldType) {
        if (fieldType == null) {
            return NULL_CONVERTER;
        }

        switch (fieldType) {
            case DATE:
                return DATE_CONVERTER;
            case DECIMAL:
                return DECIMAL_CONVERTER;
            case BOOLEAN:
                return BOOL_CONVERTER;
            case INTEGER:
                return INTEGER_CONVERTER;
            case LONG:
                return LONG_CONVERTER;
            case LONG_TEXT:
                return LONG_TEXT_CONVERTER;
            case STRING:
            default:
                return STRING_CONVERTER;
        }
    }

    /**
     * 推测字段类型
     *
     * @param val
     * @param cell
     * @return
     */
    public static FieldType guessType(Object val, Cell cell) {
        if (DATE_CONVERTER.guessType(val, cell)) {
            return FieldType.DATE;
        }

        BigDecimal decimalVal = DECIMAL_CONVERTER.convertToDecimal(val, cell);
        if (decimalVal != null) {
            boolean generalCell = Optional.ofNullable(cell)
                    .filter(ReadCellData.class::isInstance)
                    .map(ReadCellData.class::cast)
                    .map(ReadCellData::getDataFormatData)
                    .filter(dataFormatData -> "General".equals(dataFormatData.getFormat()))
                    .isPresent();
            if (generalCell) {
                // 如果单元格是常规，抹除小数后的0
                decimalVal = decimalVal.stripTrailingZeros();
            }

            if (decimalVal.scale() <= 0) {
                if (decimalVal.compareTo(IntegerConverter.INT_MAX) <= 0 && decimalVal.compareTo(IntegerConverter.INT_MIN) >= 0) {
                    return FieldType.INTEGER;
                }
                if (decimalVal.compareTo(LongConverter.LONG_MAX) <= 0 && decimalVal.compareTo(LongConverter.LONG_MIN) >= 0) {
                    return FieldType.LONG;
                }
            }
            return FieldType.DECIMAL;
        }

        if (BOOL_CONVERTER.guessType(val, cell)) {
            return FieldType.BOOLEAN;
        }

        if (LONG_TEXT_CONVERTER.doGuessType(val, cell)) {
            return FieldType.LONG_TEXT;
        }

        return FieldType.STRING;
    }
}
