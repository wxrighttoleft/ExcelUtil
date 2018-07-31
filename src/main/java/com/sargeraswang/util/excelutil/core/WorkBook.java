package com.sargeraswang.util.excelutil.core;

import com.sargeraswang.util.excelutil.ExcelException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Excel 文件操作类
 * @authro xkj [zhuiqiuxinf@163.com]
 * @version 1.3.0
 */
public class WorkBook {

    private WorkBook(){}

    private static WorkBook instance;

    public static synchronized WorkBook getInstance() {
        if (instance == null) {
            instance = new WorkBook();
        }
        return instance;
    }

    public Workbook getWorkBook(Extension extension) {
        switch(extension) {
            case XLSX:
                return new XSSFWorkbook();
            case XLS:
            default:
                return new HSSFWorkbook();
        }
    }

    public Workbook getWorkBook(String templatePath){
        File file = new File(templatePath);
        try {
            return getWorkBook(Extension.type(templatePath), new FileInputStream(file));
        } catch (IOException e) {
            throw new ExcelException(String.format("occurred a error at get file. file path is \"%s\"", templatePath), e);
        }
    }

    public Workbook getWorkBook(Extension extension, InputStream input) {
        try {
            switch (extension) {
                case XLSX:
                    return new XSSFWorkbook(input);
                case XLS:
                default:
                    return new HSSFWorkbook(input);
            }
        } catch (IOException e) {
            throw new ExcelException("occurred a error at get file.", e);
        }
    }

}
