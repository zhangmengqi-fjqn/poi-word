package com.melon.word;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPrBase;

/**
 * @author zhaokai
 * @date 2019-11-13
 */
public class Table {

    /**
     * 表格对象
     */
    private XWPFTable table;

    public Table(XWPFTable table) {
        this.table = table;
    }

    /**
     * 给表哥设置默认样式，这个默认样式是由文档得来的，由自己定义
     *
     * @param ctStyles 文档的默认样式对象
     */
    public void setDocumentDefaultStyles(CTStyles ctStyles) {
        CTTblPr tblPr = this.table.getCTTbl().getTblPr();
        if (tblPr != null) {
            // 这里的 tblPr 如果是 null，说明这个表格原来就是什么样式都没有，所以就不用加样式了
            // 此时这里不为 null，去设置样式
            if (!tblPr.isSetTblStyle()) {
                // 没有设置关联的 style，不用继续向下执行
                return;
            }
            String styleId = tblPr.getTblStyle().getVal();
            if (StringUtils.isBlank(styleId)) {
                return;
            }
            Style style = new Style(ctStyles);
            CTStyle ctStyle = style.getByStyleId(styleId);
            while (ctStyle != null) {
                // 开始设置
                if (ctStyle.isSetTblPr()) {
                    // 这个如果没有设置表格样式，就不设置，然后找 baseOn 的样式去设置
                    setTableDefaultStyles(ctStyle.getTblPr(), tblPr);
                }
                if (!ctStyle.isSetBasedOn()) {
                    break;
                }
                String baseId = ctStyle.getBasedOn().getVal();
                if (StringUtils.isBlank(baseId)) {
                    break;
                }
                ctStyle = style.getByStyleId(baseId);
            }
            tblPr.unsetTblStyle();
        }
    }

    /**
     * 看看表格还需要哪些样式，将默认的设置给他
     *
     * @param tblPr   默认的样式
     * @param ctTblPr 表格的样式
     */
    private void setTableDefaultStyles(CTTblPrBase tblPr, CTTblPr ctTblPr) {
        // tbl ind
        if (!ctTblPr.isSetTblInd() && tblPr.isSetTblInd()) {
            ctTblPr.setTblInd(tblPr.getTblInd());
        }
        // tbl cell mar
        if (!ctTblPr.isSetTblCellMar() && tblPr.isSetTblCellMar()) {
            ctTblPr.setTblCellMar(tblPr.getTblCellMar());
        }
        // tbl w
        if (!ctTblPr.isSetTblW() && tblPr.isSetTblW()) {
            ctTblPr.setTblW(tblPr.getTblW());
        }
        // tbl look
        if (!ctTblPr.isSetTblLook() && tblPr.isSetTblLook()) {
            ctTblPr.setTblLook(tblPr.getTblLook());
        }
        // tbl borders
        if (!ctTblPr.isSetTblBorders() && tblPr.isSetTblBorders()) {
            ctTblPr.setTblBorders(tblPr.getTblBorders());
        }
    }
}
