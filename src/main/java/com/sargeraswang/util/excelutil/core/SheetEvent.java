package com.sargeraswang.util.excelutil.core;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * 表页设置接口
 * @author xkj [zhuiqiuxinf@163.com]
 * @version 1.3
 */
public interface SheetEvent {
    /**
     * <p>表页创建之后执行事件，此时可以做一些不影响填充数据的操作.</p>
     * <p>比如设置表页名</p>
     * @param sheet 创建的表页
     */
    void afterCreated(Sheet sheet);

    /**
     * <p>数据填充完之后执行的事件</p>
     * <p>此时数据已填充完毕，可对表页进行更丰富的操作</p>
     * @param sheet
     */
    void afterData(Sheet sheet);
}