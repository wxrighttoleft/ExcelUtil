package com.sargeraswang.util.excelutil.core;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格样式设置
 * @author xkj [zhuiqiuxinf@163.com]
 * @version 1.3
 */
public interface StyleSetup extends SheetEvent {
    /**
     * <p>初始化操作，此时工作薄已创建完成，可做一些单元格样式的初始化工作</p>
     * @param workbook 创建的工作薄
     */
    void init(Workbook workbook);

    /**
     * <p>单元格样式应用</p>
     * @param cell 创建的单元格
     */
    void cellApply(Cell cell);

    /**
     * <p>获取时间格式</p>
     * @return 时间格式
     */
    String getPattern();
}
