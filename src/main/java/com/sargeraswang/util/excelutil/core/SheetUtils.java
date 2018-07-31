package com.sargeraswang.util.excelutil.core;

import com.sargeraswang.util.excelutil.ReflectUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.SimpleDateFormat;
import java.util.*;

public class SheetUtils {
    private static Logger LG = LoggerFactory.getLogger(SheetUtils.class);
	
	private static SheetUtils instance;
	
	private SheetUtils() {}
	
	public static synchronized SheetUtils getInstance() {
		if (instance == null) {
			instance = new SheetUtils();
		}
		return instance;
	}
	
    /**
     * 每个sheet的写入
     *
     * @param sheet     页签
     * @param headers   表头
     * @param dataset   数据集合
     * @param rule      样式规则
     */
    public <T> void write2Sheet(Sheet sheet, LinkedHashMap<String,String> headers, Collection<T> dataset, StyleSetup rule) {
        int rowIndex = 0;

        // region --------标题行设置-------------
        if (MapUtils.isNotEmpty(headers)) {
            Row headerRow = sheet.createRow(rowIndex);
            int cellIndex = 0;
            for (Iterator<Map.Entry<String, String>> iterator = headers.entrySet().iterator(); iterator.hasNext();) {
                Map.Entry<String, String> entry = iterator.next();
                Cell cell = headerRow.createCell(cellIndex);
                setCellValue(cell, entry.getValue(), rule);
                cellIndex++;
            }
            rowIndex++;
        }
        // endregion

        // region ------------填充数据------------
        if (CollectionUtils.isNotEmpty(dataset)) {
            for (Iterator<T> iterator = dataset.iterator(); iterator.hasNext();) {
                Row dataRow = sheet.createRow(rowIndex);
                T dataItem = iterator.next();
                int cellIndex = 0;
                for (Iterator<Map.Entry<String, String>> headerIterator = headers.entrySet().iterator(); headerIterator.hasNext();) {
                    Map.Entry<String, String> headerItem = headerIterator.next();
                    Cell cell = dataRow.createCell(cellIndex);
                    if (dataItem instanceof Map) {
                        Map map = (Map)dataItem;
                        setCellValue(cell, map.get(headerItem.getKey()), rule);
                    } else {
                        setCellValue(cell, ReflectUtils.getValue(dataItem, headerItem.getKey()), rule);
                    }
                    cellIndex++;
                }
                rowIndex++;
            }
        }
        // endregion

        if (rule != null) {
            rule.afterData(sheet);
        }
    }

    public void setCellValue(Cell cell, Object value, StyleSetup rule){
        String textValue = null;
        if (value instanceof Integer) {
            int intValue = (Integer) value;
            cell.setCellValue(intValue);
        } else if (value instanceof Float) {
            float fValue = (Float) value;
            cell.setCellValue(fValue);
        } else if (value instanceof Double) {
            double dValue = (Double) value;
            cell.setCellValue(dValue);
        } else if (value instanceof Long) {
            long longValue = (Long) value;
            cell.setCellValue(longValue);
        } else if (value instanceof Boolean) {
            boolean bValue = (Boolean) value;
            cell.setCellValue(bValue);
        } else if (value instanceof Date) {
            Date date = (Date) value;
            String pattern = null;
            if (rule != null)
                pattern = rule.getPattern();
            if (StringUtils.isBlank(pattern))
				pattern = "yyyy-MM-dd";
            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
            textValue = sdf.format(date);
        } else if (value instanceof String[]) {
            String[] strArr = (String[]) value;
            String str = StringUtils.join(strArr, ",");
            cell.setCellValue(str);
        } else if (value instanceof Double[]) {
            Double[] douArr = (Double[]) value;
            cell.setCellValue(StringUtils.join(douArr, ","));
        } else {
            // 其它数据类型都当作字符串简单处理
            String empty = StringUtils.EMPTY;
            textValue = value == null ? empty : value.toString();
        }
        if (textValue != null) {
            RichTextString richString;
            if (cell instanceof HSSFCell) {
                richString = new HSSFRichTextString(textValue);
            } else {
                richString = new XSSFRichTextString(textValue);
            }
            cell.setCellValue(richString);
        }
        if (rule != null) {
			rule.cellApply(cell);
		}
        LG.debug(String.format("填充第[%d,%d]数据：%s", cell.getRowIndex(), cell.getColumnIndex(), value));
    }
}
