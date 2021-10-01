package com.github.openzonedy.excel.poi;

import java.time.format.DateTimeFormatter;

public enum TimeFormatterPattern {
    yyyyMMddHHmmss("yyyy/MM/dd HH:mm:ss"),
    yyyyMMdd("yyyy-MM-dd"),
    HHmmss("HH:mm:ss"),
    ;

    public final String pattern;
    public final DateTimeFormatter formatter;

    TimeFormatterPattern(String pattern) {
        this.pattern = pattern;
        this.formatter = DateTimeFormatter.ofPattern(pattern);
    }
}
