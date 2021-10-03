package io.github.openzonedy.excel;

import io.github.openzonedy.excel.util.CellUtil;
import io.github.openzonedy.excel.util.StringUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.function.Function;

public class ExcelReader extends ExcelBase {
    protected Map<Integer, String> indexMapping;

    public ExcelReader(InputStream inputStream, String password, int sheetIndex) {
        try {
            if (StringUtil.hasText(password)) {
                workbook = WorkbookFactory.create(inputStream, password);
            } else {
                workbook = WorkbookFactory.create(inputStream);
            }
            useSheet(sheetIndex);
        } catch (Exception e) {
            try {
                inputStream.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    public ExcelReader(InputStream inputStream, String password, int sheetIndex, Class<?> clazz) {
        this(inputStream, password, sheetIndex);
        this.setColumnMapping(ReflectUtil.getReadColumnMapping(clazz));
    }

    public ExcelReader(InputStream inputStream) {
        this(inputStream, null, 0);
    }

    public ExcelReader(InputStream inputStream, Class<?> clazz) {
        this(inputStream, null, 0, clazz);
    }

    public ExcelReader(InputStream inputStream, String password) {
        this(inputStream, password, 0);
    }

    public ExcelReader(InputStream inputStream, String password, Class<?> clazz) {
        this(inputStream, password, 0, clazz);
    }

    public void useSheet(String sheetName) {
        sheet = workbook.getSheet(sheetName);
    }

    public void useSheet(int sheetIndex) {
        sheet = workbook.getSheetAt(sheetIndex);
    }

    public List<Map<String, Object>> readLine(int columnMappingRowIndex, int readStartIndex) {
        List<Map<String, Object>> mapList = new ArrayList<>();

        Row columnMappingRow = sheet.getRow(columnMappingRowIndex);
        readColumnMappingRow(columnMappingRow);

        int lastRowNum = sheet.getLastRowNum();
        for (int i = readStartIndex; i <= lastRowNum; i++) {
            Map<String, Object> rowMap = readRow(sheet.getRow(i));
            if (Objects.nonNull(rowMap)) {
                mapList.add(rowMap);
            }
        }
        return mapList;
    }

    public <T> List<T> readLine(int columnMappingRowIndex, int readStartIndex, Class<T> clazz) {
        try {
            List<Map<String, Object>> mapList = readLine(columnMappingRowIndex, readStartIndex);
            return ReflectUtil.mapToBean(mapList, columnMapping, clazz);
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException();
        }
    }

    public <T, R> List<T> readLine(int columnMappingRowIndex, int readStartIndex, Function<String, R> modifer, Class<T> clazz) {
        try {
            List<Map<String, Object>> mapList = readLine(columnMappingRowIndex, readStartIndex);
            for (Map<String, Object> lineMap : mapList) {
                for (Map.Entry<String, Object> entry : lineMap.entrySet()) {
                    Object value = entry.getValue();
                    if (value instanceof String) {
                        value = modifer.apply((String) value);
                        lineMap.put(entry.getKey(), value);
                    }
                }
            }
            return ReflectUtil.mapToBean(mapList, columnMapping, clazz);
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException();
        }
    }


    public void readColumnMappingRow(Row row) {
        Map<Integer, String> dataMap = new HashMap<>();
        Iterator<Cell> iterator = row.cellIterator();
        while (iterator.hasNext()) {
            Cell cell = iterator.next();
            dataMap.put(cell.getColumnIndex(), cell.getStringCellValue());
        }
        this.indexMapping = dataMap;
    }

    public Row readRow(int index) {
        return sheet.getRow(index);
    }

    public Map<String, Object> readRow(Row row) {
        Map<String, Object> dataMap = new LinkedHashMap<>();
        Iterator<Cell> iterator = row.cellIterator();
        boolean notEmpty = false;
        while (iterator.hasNext()) {
            Cell cell = iterator.next();
            if (this.indexMapping.containsKey(cell.getColumnIndex())) {
                Object cellValue = CellUtil.getCellValue(cell, cell.getCellType());
                dataMap.put(this.indexMapping.get(cell.getColumnIndex()), cellValue);
                if (skipEmptyRow && (Objects.nonNull(cellValue) && !StringUtil.isEmpty(cellValue.toString()))) {
                    notEmpty = true;
                }
            }
        }
        return notEmpty ? dataMap : null;
    }

}
