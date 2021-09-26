package com.github.openzonedy.excel;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;

public class ExcelDTO {
    private Byte byteItem;
    private Short shortItem;
    private Integer intItem;
    private Long longItem;
    private Float floatItem;
    private Double doubleItem;
    private Character charItem;
    private String StringItem;
    @ExcelDesc(booleanPattern = "是/否")
    private Boolean boolItem;
    @ExcelDesc(valueSetter = "enumItemSetter", valueGetter = "enumItemGetter")
    private ExcelEnum enumItem;
    @ExcelDesc(timePattern = "yyyy//MM//dd HH:mm:ss")
    private LocalDateTime localDateTime;
    private LocalDate localDate;
    private LocalTime localTime;
    private Date date;


    public String enumItemGetter() {
        return enumItem.getDesc();
    }

    public void enumItemSetter(String desc) {
        this.enumItem = ExcelEnum.descMap.get(desc);
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

    public ExcelEnum getEnumItem() {
        return enumItem;
    }

    public void setEnumItem(ExcelEnum enumItem) {
        this.enumItem = enumItem;
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
