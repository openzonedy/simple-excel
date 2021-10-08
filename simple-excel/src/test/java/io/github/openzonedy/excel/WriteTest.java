package io.github.openzonedy.excel;

import io.github.openzonedy.excel.base.CellStyleHolder;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
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
        put("enumItem1", "枚举1");
        put("localDateTime", "日期时间");
        put("localDate", "日期");
        put("localTime", "时间");
        put("date", "Date时间");
    }};

    private Map<String, String[]> optionsMap = new HashMap<>() {{
        put("enumItem1", new String[]{"XLS", "XLSX"});
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
        dto1.setEnumItem1(ExcelEnum.XLSX);
        dto1.setEnumItem2(ExcelEnum.XLSX);
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
        dto2.setEnumItem1(ExcelEnum.XLSX);
        dto2.setEnumItem2(ExcelEnum.XLSX);
        dto2.setLocalDateTime(LocalDateTime.now());
        dto2.setLocalDate(LocalDate.now());
        dto2.setLocalTime(LocalTime.now());
        dto2.setDate(new Date());
        dataLine.add(dto1);
        dataLine.add(new ExcelDTO());
        dataLine.add(dto2);
    }

    @Test
    public void write1() throws Exception {
        ExcelWriter writer = ExcelHelper.getWriter(columnMapping, true);
        writer.writeHeadLine("测试数据");
        writer.writeHeadLine(columnMapping.values());
        writer.writeLine(dataLine, ExcelDTO.class);
        writer.autoSizeColumnAll();
        writer.writeToOutputStream(new FileOutputStream("测试数据.xlsx"));
    }

    @Test
    public void write2() throws Exception {
        ExcelWriter writer = ExcelHelper.getWriter(ExcelDTO.class, true);
        writer.writeHeadLine("测试数据");
        writer.writeHeadLine(writer.columnMapping.values());
        writer.writeLine(dataLine, ExcelDTO.class);
        writer.autoSizeColumnAll();
        writer.writeToOutputStream(new FileOutputStream("测试数据BEAN.xlsx"));
    }

    @Test
    public void write3() throws Exception {
        ExcelWriter writer = ExcelHelper.getWriter(columnMapping, true);
        writer.setSkipEmptyRow(true);
        writer.setOptionsMap(optionsMap);
        CellStyleHolder cellStyleHolder = writer.getCellStyleHolder();
        XSSFCellStyle headCellStyle = (XSSFCellStyle) cellStyleHolder.getHeadCellStyle();
        headCellStyle.setFillForegroundColor(writer.cellStyleHolder.createXSSFColor("D3D3D3"));
        writer.writeHeadLine(columnMapping.values());
        writer.writeLine(dataLine, ExcelDTO.class);
        writer.addDataValidation(1, 1000, ExcelDTO.class);
        writer.autoSizeColumnAll();
        writer.writeToOutputStream(new FileOutputStream("测试数据-自定义顺序列.xlsx"));
    }

    @Test
    public void write4() throws Exception {
        ExcelWriter writer = ExcelHelper.getWriter( true);
        CellStyleHolder cellStyleHolder = writer.getCellStyleHolder();
        XSSFCellStyle headCellStyle = (XSSFCellStyle) cellStyleHolder.getHeadCellStyle();
        headCellStyle.setFillForegroundColor(writer.cellStyleHolder.createXSSFColor("D3D3D3"));

        writer.addColumnMapping("byteItem", "字节", new String[]{"Byte"});
        writer.addColumnMapping("shortItem", "短整型", new String[]{"1","2", "3"});

        writer.writeHeadLine(writer.getColumnMapping().values());
        writer.writeLine(dataLine, ExcelDTO.class);
        writer.addDataValidation(1, 1000);
        writer.autoSizeColumnAll();
        writer.writeToOutputStream(new FileOutputStream("测试数据-自定义顺序列-2.xlsx"));
    }
}
