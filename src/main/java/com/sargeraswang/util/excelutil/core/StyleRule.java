package com.sargeraswang.util.excelutil.core;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

public interface StyleRule {

    void sheetApply(Sheet sheet);

    void cellApply(Cell cell);

    String getPattern();
}
