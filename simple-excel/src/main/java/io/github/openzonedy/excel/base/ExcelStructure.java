package io.github.openzonedy.excel.base;

import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelStructure {
    private String column;
    private String columnName;
    private Class<?> clazz;
    private CellStyle cellStyle;
    private String[] options;

    public ExcelStructure(String column, String columnName) {
        this.column = column;
        this.columnName = columnName;
    }

    public ExcelStructure(String column, String columnName, Class<?> clazz) {
        this.column = column;
        this.columnName = columnName;
        this.clazz = clazz;
    }

    public ExcelStructure(String column, String columnName, Class<?> clazz, String[] options) {
        this.column = column;
        this.columnName = columnName;
        this.clazz = clazz;
        this.options = options;
    }

    public String getColumn() {
        return column;
    }

    public void setColumn(String column) {
        this.column = column;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public String[] getOptions() {
        return options;
    }

    public void setOptions(String[] options) {
        this.options = options;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}
