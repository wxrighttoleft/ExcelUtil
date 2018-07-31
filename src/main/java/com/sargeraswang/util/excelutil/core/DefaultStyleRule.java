package com.sargeraswang.util.excelutil.core;

import org.apache.poi.ss.usermodel.*;

public class DefaultStyleRule implements StyleRule {

    private CellStyle style1;
    private CellStyle style2;

    @Override
    public void sheetApply(Sheet sheet) {
    }

    @Override
    public void cellApply(Cell cell) {
        cell.setCellStyle(getStyle(cell));
    }

    @Override
    public String getPattern() {
        return "yyyy-MM-dd";
    }

    private CellStyle getStyle(Cell cell) {
        init(cell);
        if (cell.getRowIndex() % 2 == 0)
            return style1;
        else
            return style2;
    }

    private void init(Cell cell){
        if (style1 == null) {
            style1 = cell.getSheet().getWorkbook().createCellStyle();
            style1.setAlignment(HorizontalAlignment.CENTER);
            style1.setVerticalAlignment(VerticalAlignment.CENTER);
            style1.setBorderLeft(BorderStyle.THIN);
            style1.setBorderTop(BorderStyle.THIN);
            style1.setBorderBottom(BorderStyle.THIN);
            style1.setBorderRight(BorderStyle.THIN);
            style1.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
        }
        if (style2 == null) {
            style2 = cell.getSheet().getWorkbook().createCellStyle();
            style2.setAlignment(HorizontalAlignment.CENTER);
            style2.setVerticalAlignment(VerticalAlignment.CENTER);
            style2.setBorderLeft(BorderStyle.THIN);
            style2.setBorderTop(BorderStyle.THIN);
            style2.setBorderBottom(BorderStyle.THIN);
            style2.setBorderRight(BorderStyle.THIN);
            style2.setFillBackgroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
        }
    }
}
