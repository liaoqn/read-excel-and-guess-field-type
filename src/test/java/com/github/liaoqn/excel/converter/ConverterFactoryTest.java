package com.github.liaoqn.excel.converter;

import com.alibaba.excel.metadata.data.ReadCellData;
import com.github.liaoqn.excel.enums.FieldType;
import org.junit.Assert;
import org.junit.Test;

import java.math.BigDecimal;

/**
 * 转换excel单元格类型
 *
 * @author liaoqn
 * @date 2023/1/8
 */
public class ConverterFactoryTest {

    @Test
    public void convertLongTest() {
        Assert.assertEquals(100L, ConverterFactory.getConvert(FieldType.LONG).convert("$ 100", new ReadCellData<>("100")));
        Assert.assertEquals(-9223372036854775808L, ConverterFactory.getConvert(FieldType.LONG).convert("-9223372036854775808", new ReadCellData<>("-9223372036854775808")));
    }

    @Test
    public void convertDecimalTest() {
        Assert.assertEquals(BigDecimal.valueOf(100.1D), ConverterFactory.getConvert(FieldType.DECIMAL).convert("100.1", new ReadCellData<>(BigDecimal.valueOf(100.1D))));
        Assert.assertEquals(BigDecimal.valueOf(100.1D), ConverterFactory.getConvert(FieldType.DECIMAL).convert("$ 100.1", new ReadCellData<>("100.1")));
        Assert.assertEquals(BigDecimal.valueOf(-100.1D), ConverterFactory.getConvert(FieldType.DECIMAL).convert("-100.1", new ReadCellData<>("-100.1")));
        Assert.assertEquals(BigDecimal.valueOf(1.001D), ConverterFactory.getConvert(FieldType.DECIMAL).convert("100.1%", new ReadCellData<>("100.1%")));
    }

    @Test
    public void guessTypeIntegerTest() {
        Assert.assertEquals(FieldType.INTEGER, ConverterFactory.guessType("$ 100", new ReadCellData<>(BigDecimal.valueOf(100))));
    }

    @Test
    public void guessTypeLongTest() {
        Assert.assertEquals(FieldType.LONG, ConverterFactory.guessType("-9223372036854775808", new ReadCellData<>("-9223372036854775808")));
    }

    @Test
    public void guessTypeDecimalTest() {
        Assert.assertEquals(FieldType.DECIMAL, ConverterFactory.guessType("$ 100.0", new ReadCellData<>("100.0")));
        Assert.assertEquals(FieldType.DECIMAL, ConverterFactory.guessType("100.1%", new ReadCellData<>("1.001")));
        Assert.assertEquals(FieldType.DECIMAL, ConverterFactory.guessType("-100.1", new ReadCellData<>("-100.1")));
    }

    @Test
    public void guessTypStringTest() {
        Assert.assertEquals(FieldType.STRING, ConverterFactory.guessType("A 100.0", new ReadCellData<>("A 100.0")));
        Assert.assertEquals(FieldType.STRING, ConverterFactory.guessType("100.1元", new ReadCellData<>("100.1元")));
        Assert.assertEquals(FieldType.STRING, ConverterFactory.guessType("-100.1元", new ReadCellData<>("-100.1元")));
    }
}
