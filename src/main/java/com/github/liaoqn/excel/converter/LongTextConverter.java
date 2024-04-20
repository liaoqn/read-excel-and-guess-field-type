package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.github.liaoqn.excel.ExelConstants;
import com.github.liaoqn.excel.enums.FieldType;

/**
 * 转换为长文本
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class LongTextConverter extends Converter<String> {

    @Override
    public FieldType destinationType() {
        return FieldType.LONG_TEXT;
    }

    @Override
    protected String doConvert(Object val, Cell cell) {
        if (val == null) {
            return "";
        }
        return String.valueOf(val);
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        return val != null && val instanceof String && String.valueOf(val).length() > ExelConstants.STRING_MAX_LENGTH;
    }

}
