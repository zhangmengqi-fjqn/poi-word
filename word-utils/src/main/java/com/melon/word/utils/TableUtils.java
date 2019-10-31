package com.melon.word.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;

import java.math.BigInteger;

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
        if (document.getDocument().getBody().isSetSectPr()) {
            CTSectPr sectPr = document.getDocument().getBody().getSectPr();
            CTPageSz pgSz = sectPr.getPgSz();
            CTPageMar pgMar = sectPr.getPgMar();
            BigInteger totalWidth = pgSz.getW();
            if (pgMar != null) {
                // 不为空继续计算
                totalWidth = totalWidth.subtract(pgMar.getLeft()).subtract(pgMar.getRight());
            }
            // 每个的宽度
            BigInteger perWidth = totalWidth.divide(BigInteger.valueOf(cols));
            // 最后一个的宽度
            BigInteger lastWidth = totalWidth.subtract(perWidth.multiply(BigInteger.valueOf(cols - 1)));
            for (XWPFTableRow row : table.getRows()) {
                for (int i = 0; i < row.getTableCells().size(); i++) {
                    XWPFTableCell cell = row.getCell(i);
                    // 设置宽度类型, DXA 目测应该是十进制
                    cell.setWidthType(TableWidthType.DXA);
                    if (i == row.getTableCells().size() - 1) {
                        // 最后一个单元格
                        cell.setWidth(lastWidth.toString());
                    } else {
                        // 不是最后一个单元格
                        cell.setWidth(perWidth.toString());
                    }
                }
            }
        }
        return table;
    }

    /**
     * 拷贝表格的样式
     *
     * @param oldTable 旧的表格
     * @param newTable 新的表格
     */
    public static void copyStyles(XWPFTable oldTable, XWPFTable newTable) {
        CTTblPr oldTblPr = oldTable.getCTTbl().getTblPr();
        CTTblPr newTblPr = newTable.getCTTbl().getTblPr();
        if (newTblPr == null) {
            newTable.getCTTbl().setTblPr(oldTblPr);
        }
    }
}
