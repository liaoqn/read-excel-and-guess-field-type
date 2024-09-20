package com.github.liaoqn.excel.converter;

import com.github.liaoqn.excel.enums.FieldType;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.time.FastDateFormat;
import org.junit.Assert;
import org.junit.Test;

/**
 * 转换为日期
 *
 * @author liaoqn
 * @date 2023/4/27
 */
@Slf4j
public class DateConverterTest {

    private static final FastDateFormat DATE_FORMAT = FastDateFormat.getInstance("yyyy-MM-dd HH:mm:ss:SSS");
    private final Converter<?> dateConverter = ConverterFactory.getConvert(FieldType.DATE);

    @Test
    public void doConvertTest() {
        Assert.assertEquals("2000-01-02 03:04:05:006", convert("2000年1月2日3点4分5秒6毫秒"));
    }

    @Test
    public void doConvertTest1() {
        Assert.assertEquals("2000-01-02 03:04:05:000", convert("2000年1月2日3点4分5秒"));
    }

    @Test
    public void doConvertTest2() {
        Assert.assertEquals("2000-01-02 03:04:00:000", convert("2000年1月2日3点4分"));
    }

    @Test
    public void doConvertTest3() {
        Assert.assertEquals("2000-01-02 00:00:00:000", convert("2000年1月2日"));
    }


    @Test
    public void doConvertTest4() {
        Assert.assertEquals("2000-01-02 03:04:05:006", convert("2000-1-2 3:4:5:6"));
    }

    @Test
    public void doConvertTest5() {
        Assert.assertEquals("2000-01-02 03:04:05:000", convert("2000-1-2 3点4分5秒"));
    }

    @Test
    public void doConvertTest6() {
        Assert.assertEquals("2000-01-02 03:04:00:000", convert("2000-1-2 3点4分"));
    }

    @Test
    public void doConvertTest7() {
        Assert.assertEquals("2000-01-02 00:00:00:000", convert("2000-1-2"));
    }

    @Test
    public void doConvertTest8() {
        Assert.assertEquals("2000-01-02 00:00:00:000", convert("2000/1/2"));
    }

    @Test
    public void doConvertTest9() {
        Assert.assertEquals("2000-01-02 03:04:05:006", convert("2000年01月02日03点04分05秒006毫秒"));
    }

    @Test
    public void doConvertTest10() {
        Assert.assertEquals("2020-01-01 01:01:01:000", convert("2020/01/01 01:01:01"));
    }

    @Test
    public void doConvertTest11() {
        Assert.assertEquals("2020-01-01 01:01:01:002", convert("2020/01/01 01:01:01.002"));
    }

    @Test
    public void doConvertTest12() {
        Assert.assertNull(convert("01:01:01"));
    }

    @Test
    public void doGuessTypeTest() {
        boolean success = dateConverter.guessType("2000年1月1日 23:59:59:01", null);
        Assert.assertTrue(success);
    }

    @Test
    public void doConvertTest13() {
        Assert.assertEquals("2016-01-02 22:11:00:000", convert("2016-1-2 22:11"));
    }

    private String convert(String test) {
        Object date = dateConverter.convert(test, null);
        if (date == null) {
            log.info("null");
            return null;
        } else {
            String format = DATE_FORMAT.format(date);
            log.info(format);
            return format;
        }


    }

}
