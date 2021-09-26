package com.github.openzonedy.excel;

import java.util.HashMap;
import java.util.Map;

public enum ExcelEnum {
    XLS("03版本"), XLSX("07版本");

    private final String desc;

    public static final Map<String, ExcelEnum> descMap = new HashMap<>();

    static {
        for (ExcelEnum value : values()) {
            descMap.put(value.desc, value);
        }
    }

    ExcelEnum(String desc) {
        this.desc = desc;
    }

    public String getDesc() {
        return desc;
    }
}
