package com.sargeraswang.util.excelutil.core;

public enum Extension {
    XLS,
    XLSX;

    public static Extension type(String filePath){
        if (filePath.endsWith(".xlsx"))
            return XLSX;
        else
            return XLS;
    }
}
