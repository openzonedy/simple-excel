package com.github.openzonedy.excel;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.*;
import java.util.stream.Collectors;

public class ReflectUtil {

    public static List<Map<String, Object>> beanToMap(List<?> beanList, Map<String, String> columnMapping, Class<?> clazz) {
        try {
            List<Map<String, Object>> dataList = new ArrayList<>();
            List<Field> fieldList = getAllWriteDeclaredFields(clazz, columnMapping);

            for (Object bean : beanList) {
                Map<String, Object> dataMap = new HashMap<>(columnMapping.size());
                for (Field field : fieldList) {
                    Object value = field.get(bean);
                    if (Objects.isNull(value)) {
                        dataMap.put(field.getName(), "");
                    }

                    ExcelDesc excelDesc = field.getAnnotation(ExcelDesc.class);
                    if (Objects.isNull(excelDesc)) {
                        dataMap.put(field.getName(), value);
                    } else {
                        /**
                         * 1. valueGetter取值
                         */
                        String valueGetter = excelDesc.valueGetter();
                        if (valueGetter != null && valueGetter.length() > 0) {
                            Optional<Method> any = Arrays.stream(clazz.getDeclaredMethods()).filter(n -> valueGetter.equals(n.getName())).findAny();
                            if (any.isPresent()) {
                                Method getter = any.get();
                                Object invoke = getter.invoke(bean);
                                dataMap.put(field.getName(), invoke);
                                continue;
                            }
                        }

                        /**
                         * 2. timePattern格式化时间
                         */
                        String timePattern = excelDesc.timePattern();
                        if (timePattern != null && timePattern.length() > 0) {
                            String formatValue = "";
                            if (value instanceof TemporalAccessor) {
                                formatValue = DateTimeFormatter.ofPattern(timePattern).format((TemporalAccessor) value);
                                dataMap.put(field.getName(), formatValue);
                                continue;
                            } else if (value instanceof Date) {
                                LocalDateTime dateTime = LocalDateTime.ofInstant(((Date) value).toInstant(), ZoneId.systemDefault());
                                formatValue = dateTime.format(DateTimeFormatter.ofPattern(timePattern));
                                dataMap.put(field.getName(), formatValue);
                                continue;
                            }
                        }

                        /**
                         * booleanPattern格式化布尔值
                         */
                        String booleanPattern = excelDesc.booleanPattern();
                        if (booleanPattern != null && booleanPattern.length() > 0) {
                            String[] split = booleanPattern.strip().split("/");
                            String booleanValue = ((boolean) value) ? split[0] : split[1];
                            dataMap.put(field.getName(), booleanValue);
                            continue;
                        }
                    }
                }
                dataList.add(dataMap);
            }
            return dataList;
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }


    public static <T> List<T> mapToBean(List<Map<String, Object>> mapList, Map<String, String> columnMapping, Class<?> clazz) {
        try {
            Map<String, Field> fieldMap = getAllReadDeclaredFields(clazz, columnMapping);
            List<T> dataList = new ArrayList<>();
            for (Map<String, Object> dataMap : mapList) {
                Object bean = clazz.getDeclaredConstructor().newInstance();
                for (Map.Entry<String, Object> entry : dataMap.entrySet()) {
                    String columnName = entry.getKey();
                    Object value = entry.getValue();
                    if (Objects.isNull(value)) {
                        continue;
                    }

                    Field field = fieldMap.get(columnName);

                    ExcelDesc excelDesc = field.getAnnotation(ExcelDesc.class);
                    if (Objects.isNull(excelDesc)) {
                        field.set(bean, field.getType().cast(value));
                    } else {
                        /**
                         * 1. valueGetter取值
                         */
                        String valueSetter = excelDesc.valueSetter();
                        if (valueSetter != null && valueSetter.length() > 0) {
                            Optional<Method> any = Arrays.stream(clazz.getDeclaredMethods()).filter(n -> valueSetter.equals(n.getName())).findAny();
                            if (any.isPresent()) {
                                Method setter = any.get();
                                Class<?> parameterType = setter.getParameterTypes()[0];
                                setter.invoke(bean, parameterType.cast(value));
                                continue;
                            }
                        }

                        /**
                         * 2. timePattern格式化时间
                         */
                        String timePattern = excelDesc.timePattern();
                        if (timePattern != null && timePattern.length() > 0) {
                            String formatValue = "";
                            if (value instanceof TemporalAccessor) {
                                formatValue = DateTimeFormatter.ofPattern(timePattern).format((TemporalAccessor) value);
                                dataMap.put(field.getName(), formatValue);
                                continue;
                            } else if (value instanceof Date) {
                                LocalDateTime dateTime = LocalDateTime.ofInstant(((Date) value).toInstant(), ZoneId.systemDefault());
                                formatValue = dateTime.format(DateTimeFormatter.ofPattern(timePattern));
                                dataMap.put(field.getName(), formatValue);
                                continue;
                            }
                        }

                        /**
                         * booleanPattern格式化布尔值
                         */
                        String booleanPattern = excelDesc.booleanPattern();
                        if (booleanPattern != null && booleanPattern.length() > 0) {
                            String[] split = booleanPattern.strip().split("/");
                            boolean flag = value.toString().equalsIgnoreCase(split[0]);
                            field.set(bean, flag);
                            continue;
                        }
                    }
                }
            }

            return dataList;
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }


    public static List<Field> getAllWriteDeclaredFields(Class<?> clazz, Map<String, String> columnMapping) {
        List<Field> fields = new ArrayList<>();
        while (Objects.nonNull(clazz)) {
            Field[] declaredFields = clazz.getDeclaredFields();
            for (Field declaredField : declaredFields) {
                if (columnMapping.containsKey(declaredField.getName())) {
                    declaredField.setAccessible(true);
                    fields.add(declaredField);
                }
            }
            clazz = clazz.getSuperclass();
        }
        return fields;
    }

    public static Map<String, Field> getAllReadDeclaredFields(Class<?> clazz, Map<String, String> columnMapping) {
        Map<String, Field> fieldMap = new HashMap<>();


        while (Objects.nonNull(clazz)) {
            Map<String, String> colToColNameMap = columnMapping.entrySet().stream().collect(Collectors.toMap(Map.Entry::getValue, Map.Entry::getKey));
            Field[] declaredFields = clazz.getDeclaredFields();
            for (Field declaredField : declaredFields) {
                if (colToColNameMap.containsKey(declaredField.getName())) {
                    declaredField.setAccessible(true);
                    fieldMap.put(colToColNameMap.get(declaredField.getName()), declaredField);
                }
            }
            clazz = clazz.getSuperclass();
        }
        return fieldMap;
    }
}
