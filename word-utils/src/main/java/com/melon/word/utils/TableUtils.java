package com.melon.word.utils;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * 表格的工具类
 *
 * @author zhaokai
 * @date 2019-10-23
 * @see org.apache.poi.xwpf.usermodel.XWPFTable
 */
public class TableUtils {

    private TableUtils() {
    }

    /**
     * 向文档中插入一个表格
     *
     * @param document 文档对象
     * @return 表格对象
     */
    public static XWPFTable appendTable(XWPFDocument document, int rows, int cols) {
        XWPFTable table = document.createTable(rows, cols);
        XWPFStyle style = document.getStyles().getStyleWithName("table");
        return null;
    }
}
