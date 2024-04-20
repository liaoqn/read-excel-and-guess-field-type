package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.github.liaoqn.excel.enums.FieldType;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import java.text.ParseException;

/**
 * 转换为布尔
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class BoolConverter extends Converter<Boolean> {

    private final static String[] FALSE_STR = new String[]{"false", "否", "0"};
    private final static String[] TRUE_STR = new String[]{"true", "是", "1"};

    @Override
    public FieldType destinationType() {
        return FieldType.BOOLEAN;
    }

    @Override
    protected Boolean doConvert(Object val, Cell cell) throws ParseException {
        if (val == null) {
            return null;
        }

        String test = StringUtils.trim(String.valueOf(val)).toLowerCase();
        if (ArrayUtils.contains(FALSE_STR, test)) {
            return false;
        } else if (ArrayUtils.contains(TRUE_STR, test)) {
            return true;
        } else {
            throw new ParseException(test, 0);
        }
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        if (!(val instanceof String)) {
            return false;
        }

        String test = StringUtils.trim(String.valueOf(val)).toLowerCase();
        if (test.length() < 1 || test.length() > 5) {
            return false;
        }

        return ArrayUtils.contains(FALSE_STR, test) || ArrayUtils.contains(TRUE_STR, test);
    }
}
