package com.sargeraswang.util.excelutil;

import com.sargeraswang.util.excelutil.core.StyleSetup;
import org.apache.poi.ss.usermodel.*;

public class DefaultStyleRule implements StyleSetup {

    private CellStyle headerStyle;
    private CellStyle style1;
    private CellStyle style2;

    @Override
    public void init(Workbook workbook) {
        headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        style1 = workbook.createCellStyle();
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2 = workbook.createCellStyle();
        style2.cloneStyleFrom(style1);
        style2.setFillForegroundColor(IndexedColors.BROWN.getIndex());
    }

    @Override
    public void afterCreated(Sheet sheet) {

    }

    @Override
    public void afterData(Sheet sheet) {
        int cellNums = sheet.getRow(0).getPhysicalNumberOfCells();
        for (int i = 0; i < cellNums; i++) {
            sheet.autoSizeColumn(i);
        }
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
        if (cell.getRowIndex() == 0)
            return headerStyle;
        if (cell.getRowIndex() % 2 == 0)
            return style1;
        else
            return style2;
    }
}
