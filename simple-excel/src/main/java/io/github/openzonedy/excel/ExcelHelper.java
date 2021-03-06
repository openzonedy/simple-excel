package io.github.openzonedy.excel;

import java.io.InputStream;
import java.util.Map;

public class ExcelHelper {

    public static ExcelReader getReader(InputStream inputStream) {
        return new ExcelReader(inputStream);
    }

    public static ExcelReader getReader(InputStream inputStream, Class<?> clazz) {
        return new ExcelReader(inputStream, clazz);
    }

    public static ExcelWriter getWriter(boolean xssf) {
        return new ExcelWriter(xssf);
    }

    public static ExcelWriter getWriter(Map<String, String> columnMapping, boolean xssf) {
        return new ExcelWriter(columnMapping, xssf);
    }

    public static ExcelWriter getWriter(Class<?> clazz, boolean xssf) {
        return new ExcelWriter(clazz, xssf);
    }

}
