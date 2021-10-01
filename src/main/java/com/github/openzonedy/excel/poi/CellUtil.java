package com.github.openzonedy.excel.poi;

import com.github.openzonedy.excel.CellStyleHolder;
import com.github.openzonedy.excel.util.StringUtil;
import org.apache.poi.ss.usermodel.*;

import java.time.*;
import java.time.temporal.TemporalAccessor;
import java.util.Calendar;
import java.util.Date;

public class CellUtil {

    /**
     * 获取单元格值<br>
     * 如果单元格值为数字格式，则判断其格式中是否有小数部分，无则返回Long类型，否则返回Double类型
     *
     * @param cell     {@link Cell}单元格
     * @param cellType 单元格值类型{@link CellType}枚举，如果为{@code null}默认使用cell的类型
     * @return 值，类型可能为：Date、Double、Boolean、String
     */
    public static Object getCellValue(Cell cell, CellType cellType) {
        if (null == cell) {
            return null;
        }

        Object value;
        switch (cellType) {
            case NUMERIC:
                value = getNumber(cell.toString());
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case FORMULA:
                // 遇到公式时查找公式结果类型
                //暂时不实现
                value = cell.toString();
                break;
            case BLANK:
                value = "";
                break;
            case ERROR:
                final FormulaError error = FormulaError.forInt(cell.getErrorCellValue());
                value = (null == error) ? "" : error.getString();
                break;
            default:
                value = cell.toString();
        }

        return value;
    }


    /**
     * 设置单元格值<br>
     * 全部都使用传入的样式
     */
    public static void setCellValue(Cell cell, Object value, CellStyle style) {
        assert cell != null;

        if (null != style) {
            cell.setCellStyle(style);
        }
        setCellValue(cell, value);
    }

    /**
     * 自动匹配样式
     */
    public static void setCellValue(Cell cell, Object value, CellStyleHolder cellStyleHolder) {
        assert cell != null;
        //根据value的类型，自动匹配样式
        if (value instanceof Number) {
            cell.setCellStyle(cellStyleHolder.getNumberCellStyle());
        } else {
            cell.setCellStyle(cellStyleHolder.getDefaultCellStyle());
        }

        setCellValue(cell, value);
    }

    /**
     * 设置值
     *
     * @param cell
     * @param value
     */
    public static void setCellValue(Cell cell, Object value) {
        if (null == value) {
            cell.setCellValue("");
        } else if (value instanceof Date) {
            ((Date) value).toInstant();
            LocalDateTime time = LocalDateTime.ofInstant(((Date) value).toInstant(), ZoneId.systemDefault());
            String formatValue = TimeFormatterPattern.yyyyMMddHHmmss.formatter.format(time);
            cell.setCellValue(formatValue);
        } else if (value instanceof TemporalAccessor) {
            String formatValue = "";
            if (value instanceof Instant) {
                formatValue = TimeFormatterPattern.yyyyMMddHHmmss.formatter.format((TemporalAccessor) value);
            } else if (value instanceof LocalDateTime) {
                formatValue = TimeFormatterPattern.yyyyMMddHHmmss.formatter.format((TemporalAccessor) value);
            } else if (value instanceof LocalDate) {
                formatValue = TimeFormatterPattern.yyyyMMdd.formatter.format((TemporalAccessor) value);
            } else if (value instanceof LocalTime) {
                formatValue = TimeFormatterPattern.HHmmss.formatter.format((TemporalAccessor) value);
            }
            cell.setCellValue(formatValue);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        } else if (value instanceof Number) {
            cell.setCellValue(Double.parseDouble(value.toString()));
        } else {
            cell.setCellValue(value.toString());
        }
    }


    public static Object getNumber(String number) {
        Object value = null;
        if (StringUtil.hasText(number)) {
            value = getIntegerLikeNumber(number);
            if (value == null) {
                value = getFloatLikeNumber(number);
            }
        }
        return value;
    }

    /**
     * 是否时整数类型的数字，Byte,Short,Integer,Long, BigInteger
     *
     * @param number
     * @return
     */
    public static Long getIntegerLikeNumber(String number) {
        String[] split = number.split("\\.");
        if (split.length == 1) {
            return Long.parseLong(number);
        }
        if (split.length == 2) {
            if (0L == Long.parseLong(split[1])) {
                return Long.parseLong(split[0]);
            }
        }
        return null;
    }

    /**
     * 是否时浮点型数字，Float, Double,BigDemical
     *
     * @param number
     * @return
     */
    public static double getFloatLikeNumber(String number) {
        return Double.parseDouble(number);
    }

}
