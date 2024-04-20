package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.github.liaoqn.excel.ExelConstants;
import com.github.liaoqn.excel.enums.FieldType;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.FastDateFormat;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.regex.Matcher;

/**
 * 转换为String
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class StringConverter extends Converter<String> {

    private final static BigDecimal DAY_MILLISECOND = BigDecimal.valueOf(24 * 60 * 60 * 1000);

    @Override
    public FieldType destinationType() {
        return FieldType.STRING;
    }

    @Override
    protected String doConvert(Object val, Cell cell) {
        if (val == null) {
            return "";
        }
        String str = String.valueOf(val);
        if (str.length() > ExelConstants.STRING_MAX_LENGTH) {
            return str.substring(0, ExelConstants.STRING_MAX_LENGTH);
        }

        ReadCellData readCellData = getReadCellData(cell);
        if (isTimeCell(val, readCellData)) {
            return getTimeString(str, readCellData);
        }

        return str;
    }

    /**
     * Excel 得时间类型有可能被解析为数字，这里特殊处理时间类型
     *
     * @param strVal
     * @param readCellData
     * @return
     */
    private String getTimeString(String strVal, ReadCellData readCellData) {
        String dataFormat = readCellData.getDataFormatData().getFormat().replaceAll("\"", "");
        Matcher matcher = TIME_FORMAT_PATTERN.matcher(dataFormat);
        if (!matcher.find()) {
            return strVal;
        }
        String format = matcher.group(1);
        String hourFormat = matcher.group(2);
        if (StringUtils.isEmpty(format)) {
            return strVal;
        }
        format = format.replaceFirst(hourFormat, "HH");

        Date dateFrom1900 = getDateFrom1900(readCellData.getOriginalNumberValue());
        return FastDateFormat.getInstance(format).format(dateFrom1900);
    }

    /**
     * 1900日期系统:excel将1900年1月1日保存为序列号2，1900年1月2日保存为序列号3， 1900年1月3日保存为序列号 4 …… 依此类推。
     * 注意,此计算在秒及其以下单位有误差
     *
     * @param days
     * @return
     */
    private Date getDateFrom1900(BigDecimal days) {
        GregorianCalendar calendar = new GregorianCalendar(1900, Calendar.JANUARY, 1, 0, 0, 0);
        int day = days.intValue();// 小数点前
        calendar.add(Calendar.DAY_OF_MONTH, day - 2);
        BigDecimal digit = days.remainder(BigDecimal.ONE);
        if (digit.compareTo(BigDecimal.ZERO) > 0) {
            int ms = DAY_MILLISECOND.multiply(digit).setScale(0, RoundingMode.HALF_UP).intValue();
            calendar.add(Calendar.MILLISECOND, ms);
        }
        return calendar.getTime();
    }

    /**
     * 1904日期系统:excel将1904年1月1日保存为序列号0，将1904年1月2 日保存为序列号1，将 1904年1月3日保存为序列号 2 …… 依此类推。
     *
     * @param days
     * @return
     */
    private Date getDateFrom1904(BigDecimal days) {
        GregorianCalendar calendar = new GregorianCalendar(1904, Calendar.JANUARY, 1, 0, 0, 0);
        BigDecimal digit = days.remainder(BigDecimal.ONE);
        if (digit.compareTo(BigDecimal.ZERO) > 0) {
            int ms = DAY_MILLISECOND.multiply(digit).intValue();
            calendar.add(Calendar.MILLISECOND, ms);
        }
        return calendar.getTime();
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        return true;
    }

}
