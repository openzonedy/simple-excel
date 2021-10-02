package io.github.openzonedy.excel.converter;

import org.apache.poi.ss.usermodel.Cell;

public interface SimpleExcelConverter {

    Object readConvert(Cell cell);

    Cell writeConvert(Object obj);
}
