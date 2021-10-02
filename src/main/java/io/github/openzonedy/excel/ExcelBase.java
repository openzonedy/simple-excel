package io.github.openzonedy.excel;

import io.github.openzonedy.excel.base.ExcelStructure;
import io.github.openzonedy.excel.converter.EnumConverter;
import io.github.openzonedy.excel.converter.SimpleExcelConverter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
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
    protected Map<String, ExcelStructure> structureMapping = new ConcurrentHashMap<>();
    protected Map<Class<?>, SimpleExcelConverter> converterMap = new ConcurrentHashMap<>();

    protected List<ExcelStructure> excelStructureList = new ArrayList<>();

    protected static final String DEFAULT_SHEET_NAME = "Sheet1";

    {
        converterMap.put(Enum.class, new EnumConverter());
    }

    public void addColumn(String column, String columnName, Class<?> clazz, CellStyle cellStyle, String[] options) {
    }

    public void setColumnMapping(Map<String, String> columnMapping) {
        this.columnMapping.clear();
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
