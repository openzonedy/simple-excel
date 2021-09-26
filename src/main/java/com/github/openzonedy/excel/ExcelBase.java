package com.github.openzonedy.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.core.excel.converter.EnumConverter;
import org.core.excel.converter.SimpleExcelConverter;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

public abstract class ExcelBase {
    protected Workbook workbook;
    protected Sheet sheet;
    protected AtomicInteger nextRowNum = new AtomicInteger(0);
    /**
     * column -> alias  属性->名称
     */
    protected Map<String, String> columnMapping = new ConcurrentHashMap<>();
    protected Map<Class<?>, SimpleExcelConverter> converterMap = new ConcurrentHashMap<>();

    {
        converterMap.put(Enum.class, new EnumConverter());
    }

    public void setColumnMapping(Map<String, String> columnMapping) {
        this.columnMapping.putAll(columnMapping);
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Map<String, String> getColumnMapping() {
        return columnMapping;
    }
}