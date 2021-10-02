package io.github.openzonedy.excel;

import io.github.openzonedy.excel.annotation.ExcelDesc;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;

public class ExcelDTO {
    @ExcelDesc(value = "字节")
    private Byte byteItem;
    @ExcelDesc(value = "短整型")
    private Short shortItem;
    @ExcelDesc(value = "整型")
    private Integer intItem;
    @ExcelDesc(value = "长整型")
    private Long longItem;
    @ExcelDesc(value = "浮点型")
    private Float floatItem;
    @ExcelDesc(value = "双精度浮点型")
    private Double doubleItem;
    @ExcelDesc(value = "字符")
    private Character charItem;
    @ExcelDesc(value = "字符串")
    private String StringItem;
    @ExcelDesc(value = "布尔", booleanPattern = "是/否", options = {"Y", "N"})
    private Boolean boolItem;
    @ExcelDesc(value = "枚举1",valueSetter = "enumItemSetter1", valueGetter = "enumItemGetter1", options = {"XLS", "XLSX"})
    private ExcelEnum enumItem1;
    @ExcelDesc(value = "枚举2", enumValue = "desc", valueSetter = "enumItemSetter2")
    private ExcelEnum enumItem2;
    @ExcelDesc(value = "日期时间",timePattern = "yyyy//MM//dd HH:mm:ss")
    private LocalDateTime localDateTime;
    @ExcelDesc(value = "日期")
    private LocalDate localDate;
    @ExcelDesc(value = "时间")
    private LocalTime localTime;
    @ExcelDesc(value = "Date时间")
    private Date date;


    public String enumItemGetter1() {
        return enumItem1.getDesc();
    }

    public void enumItemSetter1(String desc) {
        this.enumItem1 = ExcelEnum.descMap.get(desc);
    }

    public void enumItemSetter2(String desc) {
        this.enumItem2 = ExcelEnum.descMap.get(desc);
    }

    public Integer getIntItem() {
        return intItem;
    }

    public void setIntItem(Integer intItem) {
        this.intItem = intItem;
    }

    public Short getShortItem() {
        return shortItem;
    }

    public void setShortItem(Short shortItem) {
        this.shortItem = shortItem;
    }

    public Byte getByteItem() {
        return byteItem;
    }

    public void setByteItem(Byte byteItem) {
        this.byteItem = byteItem;
    }

    public Character getCharItem() {
        return charItem;
    }

    public void setCharItem(Character charItem) {
        this.charItem = charItem;
    }

    public Long getLongItem() {
        return longItem;
    }

    public void setLongItem(Long longItem) {
        this.longItem = longItem;
    }

    public Float getFloatItem() {
        return floatItem;
    }

    public void setFloatItem(Float floatItem) {
        this.floatItem = floatItem;
    }

    public Double getDoubleItem() {
        return doubleItem;
    }

    public void setDoubleItem(Double doubleItem) {
        this.doubleItem = doubleItem;
    }

    public String getStringItem() {
        return StringItem;
    }

    public void setStringItem(String stringItem) {
        StringItem = stringItem;
    }

    public Boolean getBoolItem() {
        return boolItem;
    }

    public void setBoolItem(Boolean boolItem) {
        this.boolItem = boolItem;
    }

    public ExcelEnum getEnumItem1() {
        return enumItem1;
    }

    public void setEnumItem1(ExcelEnum enumItem1) {
        this.enumItem1 = enumItem1;
    }

    public ExcelEnum getEnumItem2() {
        return enumItem2;
    }

    public void setEnumItem2(ExcelEnum enumItem2) {
        this.enumItem2 = enumItem2;
    }

    public LocalDateTime getLocalDateTime() {
        return localDateTime;
    }

    public void setLocalDateTime(LocalDateTime localDateTime) {
        this.localDateTime = localDateTime;
    }

    public LocalDate getLocalDate() {
        return localDate;
    }

    public void setLocalDate(LocalDate localDate) {
        this.localDate = localDate;
    }

    public LocalTime getLocalTime() {
        return localTime;
    }

    public void setLocalTime(LocalTime localTime) {
        this.localTime = localTime;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}
