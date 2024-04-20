package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.github.liaoqn.excel.enums.FieldType;

/**
 * 适配null
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class NullConverter extends Converter<String> {
    @Override
    public FieldType destinationType() {
        return null;
    }

    @Override
    protected String doConvert(Object val, Cell cell) {
        return null;
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        return false;
    }

}
