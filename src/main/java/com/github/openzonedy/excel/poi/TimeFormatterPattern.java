package com.github.openzonedy.excel.poi;

import java.time.format.DateTimeFormatter;

public enum TimeFormatterPattern {
    yyyyMMddHHmmss("yyyy/MM/dd HH:mm:ss"),
    yyyyMMdd("yyyy-MM-dd"),
    HHmmss("HH:mm:ss"),
    ;

    public final DateTimeFormatter formatter;

    TimeFormatterPattern(String pattern) {
        this.formatter = DateTimeFormatter.ofPattern(pattern);
    }
}
