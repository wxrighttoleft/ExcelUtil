package com.sargeraswang.util.excelutil.core;

import com.sargeraswang.util.excelutil.ExcelException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
public class WorkBookUtils {

    private WorkBookUtils(){}

    private static WorkBookUtils instance;

    public static synchronized WorkBookUtils getInstance() {
        if (instance == null) {
            instance = new WorkBookUtils();
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
            FileInputStream fis = new FileInputStream(file);
            return WorkbookFactory.create(fis);
        } catch (IOException e) {
            throw new ExcelException(String.format("occurred a error at get file. file path is \"%s\", Please checked path of files is correct or not.", templatePath), e);
        } catch (InvalidFormatException e) {
            throw new ExcelException(String.format("occurred a error at get file. file path is \"%s\", Please checked can normal open for files.", templatePath), e);
        }
    }

    public Workbook getWorkBook(InputStream input) {
        try {
            return WorkbookFactory.create(input);
        } catch (IOException e) {
            throw new ExcelException("occurred a error at get file.", e);
        } catch (InvalidFormatException e) {
            throw new ExcelException("occurred a error at get file.", e);
        }
    }

}
