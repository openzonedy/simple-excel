package io.github.openzonedy.excel;

import io.github.openzonedy.excel.util.StringUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;

public abstract class ExcelBase {
    protected Workbook workbook;
    protected Sheet sheet;
    protected AtomicInteger nextRowNum = new AtomicInteger(0);
    protected boolean skipEmptyRow = true;
    /**
     * column -> alias  属性->名称
     */
    protected Map<String, String> columnMapping = new LinkedHashMap<>();
    protected Map<String, String[]> optionsMap = new LinkedHashMap<>();
    protected static final String DEFAULT_SHEET_NAME = "Sheet1";

    public void addColumnMapping(String column, String alias, String[] options) {
        this.columnMapping.put(column, alias);
        this.optionsMap.put(column, options);
    }

    public void addColumnMapping(String column, String alias) {
        this.columnMapping.put(column, alias);
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

    public Map<String, String[]> getOptionsMap() {
        return optionsMap;
    }

    public void setOptionsMap(Map<String, String[]> optionsMap) {
        this.optionsMap.clear();
        this.optionsMap.putAll(optionsMap);
    }

    public boolean isSkipEmptyRow() {
        return skipEmptyRow;
    }

    public void skipEmptyRow(boolean skipEmptyRow) {
        this.skipEmptyRow = skipEmptyRow;
    }

    public boolean isEmptyColumn(Object col) {
        if (Objects.nonNull(col) && StringUtil.hasText(col.toString())) {
            return Boolean.FALSE;
        }
        return Boolean.TRUE;
    }
}
