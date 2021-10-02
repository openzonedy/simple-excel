package io.github.openzonedy.excel;

import io.github.openzonedy.excel.annotation.ExcelDesc;
import io.github.openzonedy.excel.poi.TimeFormatterPattern;
import io.github.openzonedy.excel.util.StringUtil;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.*;
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
                        if (StringUtil.hasText(valueGetter)) {
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
                        if (StringUtil.hasText(timePattern)) {
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
                        if (value instanceof Boolean && StringUtil.hasText(booleanPattern)) {
                            String[] split = booleanPattern.strip().split("/");
                            String booleanValue = ((boolean) value) ? split[0] : split[1];
                            dataMap.put(field.getName(), booleanValue);
                            continue;
                        }

                        /**
                         * 枚举值取值
                         */
                        String enumValue = excelDesc.enumValue();
                        if (value.getClass().isEnum() && StringUtil.hasText(enumValue)) {
                            Field enumValueField = field.getType().getDeclaredField(enumValue);
                            enumValueField.setAccessible(true);
                            dataMap.put(field.getName(), enumValueField.get(value));
                            continue;
                        }
                        dataMap.put(field.getName(), value);
                    }
                }
                dataList.add(dataMap);
            }
            return dataList;
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }


    @SuppressWarnings("unchecked")
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
                        if (StringUtil.hasText(valueSetter)) {
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
                        if (TemporalAccessor.class.isAssignableFrom(field.getType()) || Date.class.equals(field.getType())) {
                            Optional<String> timePattern = Optional.ofNullable(StringUtil.hasText(excelDesc.timePattern()) ? excelDesc.timePattern() : null);
                            Object time = null;
                            if (TemporalAccessor.class.isAssignableFrom(field.getType())) {
                                if (LocalDateTime.class.equals(field.getType())) {
                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(timePattern.orElse(TimeFormatterPattern.yyyyMMddHHmmss.pattern));
                                    time = LocalDateTime.parse(value.toString(), formatter);
                                } else if (LocalDate.class.equals(field.getType())) {
                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(timePattern.orElse(TimeFormatterPattern.yyyyMMdd.pattern));
                                    time = LocalDate.parse(value.toString(), formatter);
                                } else if (LocalTime.class.equals(field.getType())) {
                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(timePattern.orElse(TimeFormatterPattern.HHmmss.pattern));
                                    time = LocalTime.parse(value.toString(), formatter);
                                } else if (Instant.class.equals(field.getType())) {
                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(timePattern.orElse(TimeFormatterPattern.yyyyMMddHHmmss.pattern));
                                    time = LocalDateTime.now().toInstant(OffsetDateTime.now().getOffset());
                                }
                            } else if (Date.class.equals(field.getType())) {
                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern(timePattern.orElse(TimeFormatterPattern.yyyyMMddHHmmss.pattern));
                                time = Date.from(LocalDateTime.now().toInstant(OffsetDateTime.now().getOffset()));
                            }
                            field.set(bean, time);
                        }

                        /**
                         * booleanPattern格式化布尔值
                         */
                        String booleanPattern = excelDesc.booleanPattern();
                        if (field.getType().equals(Boolean.class) && StringUtil.hasText(booleanPattern)) {
                            String[] split = booleanPattern.strip().split("/");
                            boolean flag = value.toString().equalsIgnoreCase(split[0]);
                            field.set(bean, flag);
                            continue;
                        }

                        if (field.getType().isEnum()) {
                            Method valueOf = field.getType().getDeclaredMethod("valueOf", String.class);
                            Object invoke = valueOf.invoke(bean, value.toString());
                            field.set(bean, invoke);
                        }

                        if (Number.class.isAssignableFrom(field.getType())) {
                            Object number = null;
                            if (Byte.class.equals(field.getType())) {
                                number = Byte.valueOf(value.toString());
                            } else if (Short.class.equals(field.getType())) {
                                number = Short.valueOf(value.toString());
                            } else if (Integer.class.equals(field.getType())) {
                                number = Integer.valueOf(value.toString());
                            } else if (Long.class.equals(field.getType())) {
                                number = Long.valueOf(value.toString());
                            } else if (Float.class.equals(field.getType())) {
                                number = Float.valueOf(value.toString());
                            } else if (Double.class.equals(field.getType())) {
                                number = Double.valueOf(value.toString());
                            }
                            field.set(bean, number);
                        }

                        if (String.class.equals(field.getType())) {
                            field.set(bean, value.toString());
                        }

                        if (Character.class.equals(field.getType())) {
                            field.set(bean, value.toString().charAt(0));
                        }

                    }
                }
                dataList.add((T) bean);
            }
            return dataList;
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }

    public static String getColumnName(Field field) {
        return StringUtil.hasText(field.getAnnotation(ExcelDesc.class).value()) ? field.getAnnotation(ExcelDesc.class).value() : field.getName();
    }

    public static List<Field> getAllExcelDeclaredFields(Class<?> clazz) {
        List<Field> fieldList = new ArrayList<>();
        while (Objects.nonNull(clazz)) {
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                if (Objects.nonNull(field.getAnnotation(ExcelDesc.class))) {
                    field.setAccessible(true);
                    fieldList.add(field);
                }
            }
            clazz = clazz.getSuperclass();
        }
        return fieldList.stream().sorted(Comparator.comparingInt(o -> o.getAnnotation(ExcelDesc.class).order())).collect(Collectors.toList());
    }

    public static Map<String, String> getReadColumnMapping(Class<?> clazz) {
        Map<String, String> mapping = new LinkedHashMap<>();
        List<Field> fields = getAllExcelDeclaredFields(clazz);
        for (Field field : fields) {
            mapping.put(getColumnName(field), field.getName());
        }
        return mapping;
    }

    public static List<Field> getAllDeclaredFields(Class<?> clazz) {
        List<Field> fields = new ArrayList<>();
        while (Objects.nonNull(clazz)) {
            Field[] declaredFields = clazz.getDeclaredFields();
            for (Field declaredField : declaredFields) {
                declaredField.setAccessible(true);
                fields.add(declaredField);
            }
            clazz = clazz.getSuperclass();
        }
        return fields;
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
