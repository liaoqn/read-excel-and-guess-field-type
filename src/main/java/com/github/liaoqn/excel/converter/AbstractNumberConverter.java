package com.github.liaoqn.excel.converter;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 转换为数字
 *
 * @author liaoqn
 * @date 2023/4/18
 */
@Slf4j
public abstract class AbstractNumberConverter<T extends Number> extends Converter<T> {
    // 匹配货币 ￥20879.72 ，百分数 220.13%。支持科学计数法
    protected static final Pattern DECIMAL_PATTERN = Pattern.compile("^[\\p{Sc}]?[\\s]*([-+]?\\d+[.]?\\d*([Ee][+-]?\\d+)?)(%?)$");

    /**
     * 这里如果不能转换，返回null
     *
     * @param val
     * @param cell
     * @return
     */
    public BigDecimal convertToDecimal(Object val, Cell cell) {
        ReadCellData readCellData = getReadCellData(cell);
        if (isTimeCell(val, readCellData)) {
            return null;
        }

        // 数字单元格，直接返回单元格的原始值
        BigDecimal decimalVal = Optional.ofNullable(cell)
                .filter(ReadCellData.class::isInstance)
                .map(ReadCellData.class::cast)
                .filter(c -> CellDataTypeEnum.NUMBER.equals(c.getType()))
                .map(ReadCellData::getNumberValue)
                .orElse(null);
        if (decimalVal != null) {
            return decimalVal;
        }

        String test = Optional.ofNullable(val)
                .map(String::valueOf)
                .map(s -> s.replaceAll(",", ""))
                .map(StringUtils::trimToNull)
                .orElse(null);
        if (StringUtils.isEmpty(test)) {
            return null;
        }
        // 单元格的内容是否可以直接转为数字
        if (NumberUtils.isParsable(test)) {
            return new BigDecimal(test);
        }

        // 匹配货币和百分数
        Matcher matcher = DECIMAL_PATTERN.matcher(test);
        if (!matcher.find()) {
            return null;
        }
        String decimalNumStr = matcher.group(1);
        if (StringUtils.isEmpty(decimalNumStr)) {
            return null;
        }
        try {
            BigDecimal decimalNum = new BigDecimal(decimalNumStr);
            if (StringUtils.isNotEmpty(matcher.group(3))) {
                // 百分数
                decimalNum = decimalNum.divide(new BigDecimal(100));
            }
            return decimalNum;
        } catch (Exception e) {
            log.info("parse num error, decimalNumStr={}，{}", decimalNumStr, e.getMessage());
            return null;
        }
    }
}
