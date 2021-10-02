package io.github.openzonedy.excel;

import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ReadTest {

    private Map<String, String> columnMapping = new LinkedHashMap<>() {{
        put("字节", "byteItem");
        put("短整型", "shortItem");
        put("整型", "intItem");
        put("长整型", "longItem");
        put("浮点型", "floatItem");
        put("双精度浮点型", "doubleItem");
        put("字符", "charItem");
        put("字符串", "StringItem");
        put("布尔", "boolItem");
        put("枚举", "enumItem");
        put("日期时间", "localDateTime");
        put("日期", "localDate");
        put("时间", "localTime");
        put("Date时间", "date");
    }};

    @Test
    public void read1() throws Exception {
        ExcelReader reader = ExcelHelper.getReader(new FileInputStream("测试数据.xlsx"));
        List<Map<String, Object>> mapList = reader.readLine(1, 2);
        System.out.println("");
    }

    @Test
    public void read2() throws Exception {
        ExcelReader reader = ExcelHelper.getReader(new FileInputStream("测试数据BEAN.xlsx"), ExcelDTO.class);
        reader.setSkipEmptyRow(true);
        List<ExcelDTO> dtoList = reader.readLine(1, 2, ExcelDTO.class);
        System.out.println("");
    }
}
