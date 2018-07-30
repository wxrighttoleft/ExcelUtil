package com.sargeraswang.util.ExcelUtil.core;

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
