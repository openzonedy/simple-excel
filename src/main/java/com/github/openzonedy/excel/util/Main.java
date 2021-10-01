package com.github.openzonedy.excel.util;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.time.temporal.TemporalAdjuster;

public class Main {
    public static void main(String[] args) {
        String time = "2021-10-01 12:12:12";
        TemporalAccessor parse = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss").parse(time, LocalDateTime::from);
        System.out.println("");
    }
}
