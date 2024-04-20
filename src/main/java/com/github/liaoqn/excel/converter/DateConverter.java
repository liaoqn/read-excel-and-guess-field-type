package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.Cell;
import com.github.liaoqn.excel.enums.FieldType;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.FastDateFormat;

import java.text.ParseException;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 转换为日期
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public class DateConverter extends Converter<Date> {

    private static final Pattern DATE_PATTERN = Pattern.compile("^(\\d{4})[-年/](\\d{1,2})[-月/](\\d{1,2})日?\\s?((\\d{1,2})[:时点])?((\\d{1,2})[:分])?((\\d{1,2})[:秒.]?)?((\\d{1,3})(毫秒)?)?$");

    private static final FastDateFormat DATE_FORMAT = FastDateFormat.getInstance("yyyy-MM-dd HH:mm:ss:SSS");
    private static final String DATE_STRING = "%s-%s-%s %s:%s:%s:%s";

    private static final String ZERO = "0";
    private static final String TWO_ZERO = "00";
    private static final String THREE_ZERO = "000";

    @Override
    public FieldType destinationType() {
        return FieldType.DATE;
    }

    @Override
    protected Date doConvert(Object val, Cell cell) throws ParseException {
        if (val == null) {
            return null;
        }
        String dateStr = StringUtils.trimToEmpty(String.valueOf(val));
        Matcher matcher = DATE_PATTERN.matcher(dateStr);
        if (matcher.find()) {
            String year = matcher.group(1);
            String month = fillZero(matcher.group(2));
            String day = fillZero(matcher.group(3));
            String hour = fillZero(matcher.group(5));
            String min = fillZero(matcher.group(7));
            String second = fillZero(matcher.group(9));
            String miller = fillZero(matcher.group(11), 3);
            String dateFormatStr = String.format(DATE_STRING, year, month, day, hour, min, second, miller);

            return DATE_FORMAT.parse(dateFormatStr);
        }

        return null;
    }

    @Override
    protected boolean doGuessType(Object val, Cell cell) {
        if (val == null) {
            return false;
        }
        String dateStr = StringUtils.trimToEmpty(String.valueOf(val));
        Matcher matcher = DATE_PATTERN.matcher(dateStr);
        return matcher.matches();
    }

    private String fillZero(String time) {
        return fillZero(time, 2);
    }

    /**
     * 把一位的时间补0成两位
     *
     * @param time
     * @param targetLength
     * @return
     */
    private String fillZero(String time, int targetLength) {
        if (StringUtils.isEmpty(time)) {
            if (targetLength == 3) {
                return THREE_ZERO;
            } else {
                return TWO_ZERO;
            }
        }

        if (time.length() >= targetLength) {
            return time;
        }
        int zeroNum = targetLength - time.length();
        if (zeroNum == 2) {
            return TWO_ZERO + time;
        } else {
            return ZERO + time;
        }
    }
}
