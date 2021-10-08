package io.github.openzonedy.excel;

import io.github.openzonedy.excel.util.CellUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelWriter extends ExcelBase {
    protected CellStyleHolder cellStyleHolder;

    public ExcelWriter(boolean xssf) {
        try {
            /**
             * 等同于workbook = new XSSFWorkbook();
             */
            workbook = WorkbookFactory.create(xssf);
            sheet = workbook.createSheet(DEFAULT_SHEET_NAME);
            cellStyleHolder = new CellStyleHolder(workbook);
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }

    public ExcelWriter(Map<String, String> columnMapping, boolean xssf) {
        this(xssf);
        this.columnMapping = columnMapping;
    }

    public ExcelWriter(Class<?> clazz, boolean xssf) {
        this(xssf);
        List<Field> fields = ReflectUtil.getAllExcelDeclaredFields(clazz);
        this.columnMapping = fields.stream().collect(Collectors.toMap(Field::getName, ReflectUtil::getColumnName, (o1, o2) -> o1, LinkedHashMap::new));
    }

    public void writeHeadLine(Collection<String> headLine) {
        Row row = sheet.createRow(nextRowNum.getAndIncrement());
        int index = 0;
        for (String columnName : headLine) {
            Cell cell = row.createCell(index++, CellType.STRING);
            cell.setCellValue(columnName);
            cell.setCellStyle(cellStyleHolder.headCellStyle);
        }
    }


    public void writeHeadLine(String headLine) {
        Row row = sheet.createRow(nextRowNum.getAndIncrement());
        Cell cell = row.createCell(0);
        cell.setCellValue(headLine);
        cell.setCellStyle(cellStyleHolder.headCellStyle);
        merge(row.getRowNum(), row.getRowNum(), 0, columnMapping.size() - 1, cellStyleHolder.headCellStyle);
    }

    public void writeHeadLine(String headLine, int mergeStartCol, int mergeEndCol) {
        Row row = sheet.createRow(nextRowNum.getAndIncrement());
        Cell cell = row.createCell(0);
        cell.setCellValue(headLine);
        cell.setCellStyle(cellStyleHolder.headCellStyle);
        merge(row.getRowNum(), row.getRowNum(), mergeStartCol, mergeEndCol, cellStyleHolder.headCellStyle);
    }

    public void writeLine(List<?> beanList, Class<?> clazz) {
        List<Map<String, Object>> mapList = ReflectUtil.beanToMap(beanList, columnMapping, clazz);
        writeLine(mapList);
    }

    public void writeLine(List<Map<String, Object>> mapList) {
        for (Map<String, Object> dataLine : mapList) {
            if (skipEmptyRow && dataLine.values().stream().allMatch(this::isEmptyColumn)) {
                continue;
            }
            Row row = sheet.createRow(nextRowNum.getAndIncrement());
            List<String> columns = new ArrayList<>(columnMapping.keySet());
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = row.createCell(i);
                CellUtil.setCellValue(cell, dataLine.get(columns.get(i)), cellStyleHolder);
            }
        }
    }

    /**
     * 增加下拉列表
     *
     * @param regions {@link CellRangeAddressList} 指定下拉列表所占的单元格范围
     * @param options 下拉列表内容
     * @return this
     * @since 4.6.2
     */
    public void addDataValidation(CellRangeAddressList regions, String... options) {
        DataValidationHelper validationHelper = this.sheet.getDataValidationHelper();
        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);

        //设置下拉框数据
        DataValidation dataValidation = validationHelper.createValidation(constraint, regions);

        //处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }

        sheet.addValidationData(dataValidation);
    }

    public void addDataValidation(int firstRow, int lastRow, int firstCol, int lastCol, String... options) {
        this.addDataValidation(new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol), options);
    }

    public void addDataValidation(int firstRow, int lastRow, int col, String... options) {
        this.addDataValidation(new CellRangeAddressList(firstRow, lastRow, col, col), options);
    }

    public void addDataValidation(int firstRow, int lastRow, Map<String, String[]> optionsMap) {
        optionsMap.forEach((k, v) -> {
            List<String> columns = new ArrayList<>(columnMapping.keySet());
            int col = columns.indexOf(k);
            this.addDataValidation(new CellRangeAddressList(firstRow, lastRow, col, col), v);
        });
    }

    public void addDataValidation(int firstRow, int lastRow) {
        if (this.optionsMap.size() > 0) {
            this.addDataValidation(firstRow, lastRow, this.optionsMap);
        }
    }

    public void addDataValidation(int firstRow, int lastRow, Class<?> clazz) {
        this.optionsMap.clear();
        this.optionsMap.putAll(ReflectUtil.getOptionsMap(clazz));
        this.addDataValidation(firstRow, lastRow);
    }

    public void merge(int firstRow, int lastRow, int firstCol, int lastCol, CellStyle cellStyle) {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

        if (Objects.nonNull(cellStyle)) {
            RegionUtil.setBorderTop(cellStyle.getBorderTop(), cellRangeAddress, sheet);
            RegionUtil.setBorderBottom(cellStyle.getBorderBottom(), cellRangeAddress, sheet);
            RegionUtil.setBorderLeft(cellStyle.getBorderLeft(), cellRangeAddress, sheet);
            RegionUtil.setBorderRight(cellStyle.getBorderRight(), cellRangeAddress, sheet);

            RegionUtil.setTopBorderColor(cellStyle.getTopBorderColor(), cellRangeAddress, sheet);
            RegionUtil.setBottomBorderColor(cellStyle.getBottomBorderColor(), cellRangeAddress, sheet);
            RegionUtil.setLeftBorderColor(cellStyle.getRightBorderColor(), cellRangeAddress, sheet);
            RegionUtil.setRightBorderColor(cellStyle.getLeftBorderColor(), cellRangeAddress, sheet);
        }

        sheet.addMergedRegion(cellRangeAddress);
    }

    public void merge(int firstCol, int lastCol, CellStyle cellStyle) {
        merge(nextRowNum.get() - 1, nextRowNum.get() - 1, firstCol, lastCol, cellStyle);
    }

    public void autoSizeColumnAll() {
        for (int i = 0; i < columnMapping.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }


    public void writeToBrowser(HttpServletResponse response, String fileName) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8));
        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);
        workbook.close();
        out.flush();
        out.close();
    }


    public void writeToOutputStream(OutputStream out) throws IOException {
        workbook.write(out);
        workbook.close();
        out.flush();
        out.close();
    }

    public CellStyleHolder getCellStyleHolder() {
        return cellStyleHolder;
    }

    public void setCellStyleHolder(CellStyleHolder cellStyleHolder) {
        this.cellStyleHolder = cellStyleHolder;
    }
}
