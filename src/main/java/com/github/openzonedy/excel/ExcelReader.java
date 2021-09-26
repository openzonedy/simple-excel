package com.github.openzonedy.excel;

import com.github.openzonedy.excel.poi.CellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class ExcelReader extends ExcelBase {
    protected Map<Integer, String> indexMapping;

    public ExcelReader(InputStream inputStream, int sheetIndex) {
        try {
            workbook = new XSSFWorkbook(inputStream);
            useSheet(sheetIndex);
        } catch (Exception e) {
            try {
                inputStream.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    public ExcelReader(InputStream inputStream) {
        this(inputStream, 0);
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
            mapList.add(rowMap);
        }
        return mapList;
    }

    public <T> List<T> readLine(int columnMappingRowIndex, int readStartIndex, Class<T> clazz) {
        List<T> dataList = new ArrayList<>();
        try {
            List<Map<String, Object>> mapList = readLine(columnMappingRowIndex, readStartIndex);
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
        while (iterator.hasNext()) {
            Cell cell = iterator.next();
            if (this.indexMapping.containsKey(cell.getColumnIndex())) {
                dataMap.put(this.indexMapping.get(cell.getColumnIndex()), CellUtil.getCellValue(cell, cell.getCellType()));
            }
        }
        return dataMap;
    }

}
