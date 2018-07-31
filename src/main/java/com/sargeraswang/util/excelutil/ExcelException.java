package com.sargeraswang.util.excelutil;

public class ExcelException extends RuntimeException {
    public ExcelException(String s) {
        super(s);
    }

    public ExcelException(String s, Throwable throwable) {
        super(s, throwable);
    }

    public ExcelException(Throwable throwable) {
        super(throwable);
    }

    public ExcelException(String s, Throwable throwable, boolean b, boolean b1) {
        super(s, throwable, b, b1);
    }
}
