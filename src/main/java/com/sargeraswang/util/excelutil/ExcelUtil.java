package com.sargeraswang.util.excelutil;

import com.sargeraswang.util.excelutil.core.*;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * The <code>ExcelUtil</code>
 *
 * @author sargeras.wang
 * @author xkj [zhuiqiuxinf@163.com]
 * @version 1.0, Created at 2013年9月14日
 */
public class ExcelUtil {

    private static Logger LG = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 用来验证excel与Vo中的类型是否一致 <br>
     * Map<栏位类型,只能是哪些Cell类型>
     */
    private static Map<Class<?>, CellType[]> validateMap = new HashMap<>();

    static {
        validateMap.put(String[].class, new CellType[]{CellType.STRING});
        validateMap.put(Double[].class, new CellType[]{CellType.NUMERIC});
        validateMap.put(String.class, new CellType[]{CellType.STRING});
        validateMap.put(Double.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Date.class, new CellType[]{CellType.NUMERIC, CellType.STRING});
        validateMap.put(Integer.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Float.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Long.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Boolean.class, new CellType[]{CellType.BOOLEAN});
    }

    private Extension extension;

    public ExcelUtil() {
        this.extension = Extension.XLS;
    }

    public ExcelUtil(Extension extension) {
        this.extension = extension;
    }

    /**
     * 获取cell类型的文字描述
     *
     * @param cellType <pre>
     *                 CellType.BLANK
     *                 CellType.BOOLEAN
     *                 CellType.ERROR
     *                 CellType.FORMULA
     *                 CellType.NUMERIC
     *                 CellType.STRING
     *                 </pre>
     * @return
     */
    private static String getCellTypeByInt(CellType cellType) {
        if(cellType == CellType.BLANK)
            return "Null type";
        else if(cellType == CellType.BOOLEAN)
            return "Boolean type";
        else if(cellType == CellType.ERROR)
            return "Error type";
        else if(cellType == CellType.FORMULA)
            return "Formula type";
        else if(cellType == CellType.NUMERIC)
            return "Numeric type";
        else if(cellType == CellType.STRING)
            return "String type";
        else
            return "Unknown type";
    }

    /**
     * 获取单元格值
     *
     * @param cell
     * @return
     */
    private static Object getCellValue(Cell cell) {
        if (cell == null
                || (cell.getCellTypeEnum() == CellType.STRING && StringUtils.isBlank(cell
                .getStringCellValue()))) {
            return null;
        }
        CellType cellType = cell.getCellTypeEnum();
            if(cellType == CellType.BLANK)
                return null;
            else if(cellType == CellType.BOOLEAN)
                return cell.getBooleanCellValue();
            else if(cellType == CellType.ERROR)
                return cell.getErrorCellValue();
            else if(cellType == CellType.FORMULA) {
                try {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    } else {
                        return cell.getNumericCellValue();
                    }
                } catch (IllegalStateException e) {
                    return cell.getRichStringCellValue();
                }
            }
            else if(cellType == CellType.NUMERIC){
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            }
            else if(cellType == CellType.STRING)
                return cell.getStringCellValue();
            else
                return null;
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于单个sheet
     *
     * @param <T>
     * @param headers 表格属性列名数组
     * @param dataset 需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的
     *                javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     * @param out     与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     */
    public <T> void exportExcel(LinkedHashMap<String,String> headers, Collection<T> dataset, OutputStream out) {
        exportExcel(headers, dataset, out, null);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于单个sheet
     *
     * @param <T>
     * @param headers 表格属性列名数组
     * @param dataset 需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的
     *                javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     * @param out     与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     * @param rule      样式规则，个性化表格
     */
    public <T> void exportExcel(LinkedHashMap<String,String> headers, Collection<T> dataset, OutputStream out, StyleSetup rule) {
        // 声明一个工作薄
        Workbook workbook = WorkBookUtils.getInstance().getWorkBook(this.extension);
        // 生成一个表格
        Sheet sheet = workbook.createSheet();

        if (rule != null) {
            rule.init(workbook);
            rule.afterCreated(sheet);
        }

        SheetUtils.getInstance().write2Sheet(sheet, headers, dataset, rule);
        try {
            workbook.write(out);
        } catch (IOException e) {
            LG.error(e.toString(), e);
        }
    }

    public void exportExcel(String[][] datalist, OutputStream out) {
        try {
            // 声明一个工作薄
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 生成一个表格
            HSSFSheet sheet = workbook.createSheet();

            for (int i = 0; i < datalist.length; i++) {
                String[] r = datalist[i];
                HSSFRow row = sheet.createRow(i);
                for (int j = 0; j < r.length; j++) {
                    HSSFCell cell = row.createCell(j);
                    //cell max length 32767
                    if (r[j].length() > 32767) {
                        r[j] = "--此字段过长(超过32767),已被截断--" + r[j];
                        r[j] = r[j].substring(0, 32766);
                    }
                    cell.setCellValue(r[j]);
                }
            }
            //自动列宽
            if (datalist.length > 0) {
                int colcount = datalist[0].length;
                for (int i = 0; i < colcount; i++) {
                    sheet.autoSizeColumn(i);
                }
            }
            workbook.write(out);
        } catch (IOException e) {
            LG.error(e.toString(), e);
        }
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于多个sheet
     *
     * @param <T>
     * @param sheets {@link ExcelSheet}的集合
     * @param out    与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     */
    public <T> void exportExcel(List<ExcelSheet<T>> sheets, OutputStream out) {
        exportExcel(sheets, out, null);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于多个sheet
     *
     * @param <T>
     * @param sheets  {@link ExcelSheet}的集合
     * @param out     与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     */
    public <T> void exportExcel(List<ExcelSheet<T>> sheets, OutputStream out, StyleSetup rule) {
        if (CollectionUtils.isEmpty(sheets)) {
            return;
        }
        // 声明一个工作薄
        Workbook workbook = WorkBookUtils.getInstance().getWorkBook(extension);
        if (rule != null) {
            rule.init(workbook);
        }
        for (ExcelSheet<T> sheet : sheets) {
            // 生成一个表格
            Sheet s = workbook.createSheet(sheet.getSheetName());
            if (rule != null) {
                rule.afterCreated(s);
            }
            SheetUtils.getInstance().write2Sheet(s, sheet.getHeaders(), sheet.getDataset(), rule);
        }
        try {
            workbook.write(out);
        } catch (IOException e) {
            LG.error(e.toString(), e);
        }
    }

    /**
     * 把Excel的数据封装成voList
     *
     * @param clazz       vo的Class
     * @param inputStream excel输入流
     * @param pattern     如果有时间数据，设定输入格式。默认为"yyy-MM-dd"
     * @param logs        错误log集合
     * @param arrayCount  如果vo中有数组类型,那就按照index顺序,把数组应该有几个值写上.
     * @return voList
     * @throws RuntimeException
     */
    public static <T> Collection<T> importExcel(Class<T> clazz, InputStream inputStream, String pattern, ExcelLogs logs, Integer... arrayCount){
        return null;
    }

}
