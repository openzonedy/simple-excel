package com.github.openzonedy.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

public class WriteTest {

    private List<ExcelDTO> dataLine = new ArrayList<>();
    private Map<String, String> columnMapping = new LinkedHashMap<>() {{
        put("byteItem", "字节");
        put("shortItem", "短整型");
        put("intItem", "整型");
        put("longItem", "长整型");
        put("floatItem", "浮点型");
        put("doubleItem", "双精度浮点型");
        put("charItem", "字符");
        put("StringItem", "字符串");
        put("boolItem", "布尔");
        put("enumItem", "枚举");
        put("localDateTime", "日期时间");
        put("localDate", "日期");
        put("localTime", "时间");
        put("date", "Date时间");
    }};

    {
        ExcelDTO dto1 = new ExcelDTO();
        dto1.setByteItem((byte) 1);
        dto1.setShortItem((short) 5);
        dto1.setIntItem(10);
        dto1.setLongItem(20L);
        dto1.setFloatItem(1.3F);
        dto1.setDoubleItem(2.3);
        dto1.setCharItem('A');
        dto1.setStringItem("ABC");
        dto1.setBoolItem(true);
        dto1.setEnumItem(ExcelEnum.XLSX);
        dto1.setLocalDateTime(LocalDateTime.now());
        dto1.setLocalDate(LocalDate.now());
        dto1.setLocalTime(LocalTime.now());
        dto1.setDate(new Date());

        ExcelDTO dto2 = new ExcelDTO();
        dto2.setByteItem((byte) 2);
        dto2.setShortItem((short) 7);
        dto2.setIntItem(15);
        dto2.setLongItem(27L);
        dto2.setFloatItem(2.3F);
        dto2.setDoubleItem(2.0);
        dto2.setCharItem('B');
        dto2.setStringItem("ICBC");
        dto2.setBoolItem(false);
        dto2.setEnumItem(ExcelEnum.XLSX);
        dto2.setLocalDateTime(LocalDateTime.now());
        dto2.setLocalDate(LocalDate.now());
        dto2.setLocalTime(LocalTime.now());
        dto2.setDate(new Date());
        dataLine.add(dto1);
        dataLine.add(dto2);
    }

    @Test
    public void write1() throws Exception {
        ExcelWriter writer = ExcelHelper.getWriter(columnMapping);
        writer.writeHeadLine("测试数据");
        writer.writeHeadLine(columnMapping.values());
        writer.writeLine(dataLine, ExcelDTO.class);
        writer.autoSizeColumnAll();
        writer.writeToOutputStream(new FileOutputStream("测试数据.xlsx"));

    }

    @Test
    public void write2() throws Exception {
        ExcelWriter writer = ExcelHelper.getWriter(columnMapping);
        CellStyleHolder cellStyleHolder = writer.getCellStyleHolder();
        XSSFCellStyle headCellStyle = (XSSFCellStyle)cellStyleHolder.getHeadCellStyle();
        XSSFColor xssfColor = new XSSFColor(new DefaultIndexedColorMap());
        xssfColor.setARGBHex("D3D3D3");
        headCellStyle.setFillForegroundColor(xssfColor);
    }
}
