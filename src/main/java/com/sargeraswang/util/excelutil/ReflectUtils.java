package com.sargeraswang.util.excelutil;

import java.lang.reflect.Field;

public class ReflectUtils {
    public static Object getValue(Object source, String fieldName) {
        // 获取字段
        try {
            Field field = source.getClass().getDeclaredField(fieldName);
            field.setAccessible(true);
            return field.get(source);
        } catch (NoSuchFieldException e) {
            throw new ExcelException("field information not exists.", e);
        } catch (IllegalAccessException e) {
            throw new ExcelException("occurred a error at get values.", e);
        }
    }
}
