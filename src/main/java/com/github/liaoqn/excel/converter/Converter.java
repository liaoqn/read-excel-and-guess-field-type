package com.github.liaoqn.excel.converter;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.DataFormatData;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.github.liaoqn.excel.enums.FieldType;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.text.ParseException;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * excel单元格值类型转换
 *
 * @author liaoqn
 * @date 2023/4/18
 */
@Slf4j
public abstract class Converter<T> {

    protected static final Pattern TIME_FORMAT_PATTERN = Pattern.compile("^[^y]*(([hH]{1,2})[:时点]m{2}[:分]?(ss)?秒?).*$");

    /**
     * 转换后的类型
     *
     * @return
     */
    abstract FieldType destinationType();

    /**
     * 转换类型
     *
     * @param val
     * @param cell
     * @return
     */
    public T convert(Object val, Cell cell) {
        try {
            return doConvert(val, cell);
        } catch (Exception e) {
            log.warn("读取excel，转换为 {} 出错，val={}", destinationType().name(), val);
            return null;
        }
    }

    /**
     * 推测字段类型
     *
     * @param val
     * @param cell
     * @return
     */
    public boolean guessType(Object val, Cell cell) {
        try {
            if (val == null) {
                return false;
            }
            return doGuessType(val, cell);
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 转换类型
     *
     * @param val
     * @param cell
     * @return
     */
    abstract protected T doConvert(Object val, Cell cell) throws ParseException;

    abstract protected boolean doGuessType(Object val, Cell cell) throws ParseException;

    protected CellDataTypeEnum getCellType(Cell cell) {
        if (cell instanceof ReadCellData) {
            ReadCellData readCellData = (ReadCellData) cell;
            return readCellData.getType();
        }
        return null;
    }

    protected ReadCellData getReadCellData(Cell cell) {
        return Optional.ofNullable(cell)
                .filter(ReadCellData.class::isInstance)
                .map(ReadCellData.class::cast)
                .orElse(null);
    }

    protected boolean isTimeCell(Object val, ReadCellData readCellData) {
        if (readCellData == null || readCellData.getDataFormatData() == null
            || StringUtils.isEmpty(readCellData.getDataFormatData().getFormat())) {
            return false;
        }

        String dataFormat = readCellData.getDataFormatData().getFormat().replaceAll("\"", "");
        Matcher matcher = TIME_FORMAT_PATTERN.matcher(dataFormat);
        return matcher.matches();
    }
}
