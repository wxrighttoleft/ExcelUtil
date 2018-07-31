package com.sargeraswang.util.excelutil;

import org.junit.Test;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.LinkedHashMap;

public class TestExportBean {
    @Test
    public void exportXls() throws IOException {
        //用排序的Map且Map的键应与ExcelCell注解的index对应
        LinkedHashMap<String,String> map = new LinkedHashMap<>();
        map.put("a","姓名");
        map.put("b","年龄");
        map.put("c","性别");
        map.put("d","出生日期");
        Collection<Object> dataset=new ArrayList<Object>();
        dataset.add(new Model("", "", "",null));
        dataset.add(new Model(null, null, null,null));
        dataset.add(new Model("王五", "34", "男",new Date()));
        File f=new File("test.xlsx");
        OutputStream out =new FileOutputStream(f);
        
        ExcelUtil.exportExcel(map, dataset, out);

    }
}
