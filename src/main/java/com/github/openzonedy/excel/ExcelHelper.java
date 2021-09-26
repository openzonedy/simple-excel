package com.github.openzonedy.excel;

import java.io.InputStream;
import java.util.Map;

public class ExcelHelper {
    /**
     * 可以多参考Apache POI  util包 和hutool的一些写法
     *
     * @param inputStream
     * @return
     */

    public static ExcelReader getReader(InputStream inputStream) {
        return new ExcelReader(inputStream);
    }

    public static ExcelWriter getWriter() {
        return new ExcelWriter();
    }

    public static ExcelWriter getWriter(Map<String, String> columnMapping) {
        return new ExcelWriter(columnMapping);
    }

    public static void getReader() {
    }
}
