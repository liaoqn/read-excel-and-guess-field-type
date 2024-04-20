package com.github.liaoqn.excel.enums;

import java.math.BigDecimal;
import java.util.Date;

/**
 * 字段类型
 *
 * @author liaoqn
 * @date 2023/4/18
 */
public enum FieldType {
    /**
     * 文本
     */
    STRING(String.class),
    /**
     * 日期
     */
    DATE(Date.class),
    /**
     * 整数
     */
    INTEGER(Integer.class),
    /**
     * 长整数
     */
    LONG(Long.class),
    /**
     * 小数
     */
    DECIMAL(BigDecimal.class),
    /**
     * 布尔
     */
    BOOLEAN(Boolean.class),
    /**
     * 大文本，超过4000个字符
     */
    LONG_TEXT(String.class);

    /**
     * 字段对应的 Java 类型
     */
    private final Class<?> classType;

    FieldType(Class<?> classType) {
        this.classType = classType;
    }

    public Class<?> getClassType() {
        return classType;
    }
}
