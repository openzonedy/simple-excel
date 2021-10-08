package io.github.openzonedy.excel.base;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class CellStyleHolder {
    /**
     * 标题样式
     */
    public CellStyle headCellStyle;
    /**
     * 默认样式
     */
    public CellStyle defaultCellStyle;
    /**
     * 默认数字样式
     */
    public CellStyle numberCellStyle;
    /**
     * 自定义的其他样式
     */
    public Map<String, CellStyle> customCellStyleMap = new ConcurrentHashMap<>();

    public CellStyleHolder(Workbook workbook) {
        /**
         * 默认样式
         */
        CellStyle defaultCellStyle = workbook.createCellStyle();
        //垂直对齐方式
        defaultCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //水平对齐方式
        defaultCellStyle.setAlignment(HorizontalAlignment.LEFT);
        defaultCellStyle.setFont(createFontStyle(workbook, "微软雅黑", false, 12));
        this.defaultCellStyle = defaultCellStyle;

        /**
         * 头部样式
         */
        CellStyle headCellStyle = workbook.createCellStyle();
        headCellStyle.cloneStyleFrom(this.defaultCellStyle);
        headCellStyle.setFont(createFontStyle(workbook, "微软雅黑", false, 12));
        headCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headCellStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置前景色
        headCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        //设置填充方式
        headCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置边框及颜色
        setBorder(headCellStyle, BorderStyle.THIN, IndexedColors.BLACK);
        this.headCellStyle = headCellStyle;

        /**
         * 数字样式
         */
        CellStyle numberCellStyle = workbook.createCellStyle();
        numberCellStyle.cloneStyleFrom(defaultCellStyle);
        numberCellStyle.setAlignment(HorizontalAlignment.RIGHT);
        this.numberCellStyle = numberCellStyle;
    }

    /**
     * Font -> XSSFFont/HSSFFont
     *
     * @param workbook
     * @param fontName
     * @param bold
     * @param fontSize
     * @return
     */
    public Font createFontStyle(Workbook workbook, String fontName, boolean bold, int fontSize) {
        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setBold(bold);
        font.setColor(Font.COLOR_NORMAL); //IndexedColors亦可、颜色选择更多
        font.setFontHeightInPoints((short) fontSize);
        return font;
    }

    /**
     * XSSF 设置RGB颜色
     * XSSFCellStyle
     *
     * @param rgbHex
     */
    public XSSFColor createXSSFColor(String rgbHex) {
        XSSFColor xssfColor = new XSSFColor(new DefaultIndexedColorMap());
        xssfColor.setARGBHex(rgbHex);
        return xssfColor;
    }

    /**
     * 设置边框及边框颜色
     */
    private void setBorder(CellStyle cellStyle, BorderStyle borderStyle, IndexedColors color) {
        cellStyle.setBorderTop(borderStyle);
        cellStyle.setBorderBottom(borderStyle);
        cellStyle.setBorderLeft(borderStyle);
        cellStyle.setBorderRight(borderStyle);

        cellStyle.setTopBorderColor(color.index);
        cellStyle.setBottomBorderColor(color.index);
        cellStyle.setLeftBorderColor(color.index);
        cellStyle.setRightBorderColor(color.index);
    }

    public CellStyle getHeadCellStyle() {
        return headCellStyle;
    }

    public void setHeadCellStyle(CellStyle headCellStyle) {
        this.headCellStyle = headCellStyle;
    }

    public CellStyle getCellStyle() {
        return defaultCellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.defaultCellStyle = cellStyle;
    }

    public CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    public void setDefaultCellStyle(CellStyle defaultCellStyle) {
        this.defaultCellStyle = defaultCellStyle;
    }

    public CellStyle getNumberCellStyle() {
        return numberCellStyle;
    }

    public void setNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
    }

    public Map<String, CellStyle> getCustomCellStyleMap() {
        return customCellStyleMap;
    }

    public void setCustomCellStyleMap(Map<String, CellStyle> customCellStyleMap) {
        this.customCellStyleMap = customCellStyleMap;
    }
}
