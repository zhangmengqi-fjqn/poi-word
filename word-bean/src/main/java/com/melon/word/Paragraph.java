package com.melon.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.List;

/**
 * 适配器模式-段落
 *
 * @author zhaokai
 * @date 2019-11-02
 */
public class Paragraph {

    /**
     * @see XWPFParagraph
     */
    private XWPFParagraph paragraph;

    /**
     * 使用已有段落创建
     *
     * @param paragraph 段落
     */
    public Paragraph(XWPFParagraph paragraph) {
        this.paragraph = paragraph;
    }

    /**
     * 创建一个空的文档段落
     *
     * @param document 文档
     */
    public Paragraph(XWPFDocument document) {
        CTP ctp = CTP.Factory.newInstance();
        this.paragraph = new XWPFParagraph(ctp, document);
    }

    public XWPFParagraph getParagraph() {
        return paragraph;
    }

    /**
     * 给段落设置默认的段落样式
     *
     * @param ctStyles 默认的样式
     */
    public void setDocumentDefaultStyles(CTStyles ctStyles) {
        Style style = new Style(ctStyles);
        CTP ctp = paragraph.getCTP();
        if (style.getDefaultCTPPr() != null) {
            // 设置默认的段落样式
            setParagraphDefaultStyles(ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr(), style.getDefaultCTPPr());
        }
        if (style.getDefaultCTRPr() != null) {
            // 设置段落样式中的 run 的默认样式
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                CTRPr ctrPr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
                setRunDefaultStyles(ctrPr, style.getDefaultCTRPr());
            }
        }
    }

    /**
     * 给 run 样式赋值默认样式
     *
     * @param ctrPr run 的样式
     * @param rPr   默认的 run 样式
     */
    public void setRunDefaultStyles(CTRPr ctrPr, CTRPr rPr) {
        if (!ctrPr.isSetRFonts() && rPr.isSetRFonts()) {
            ctrPr.setRFonts(rPr.getRFonts());
        }
        if (!ctrPr.isSetKern() && rPr.isSetKern()) {
            ctrPr.setKern(rPr.getKern());
        }
        if (!ctrPr.isSetSz() && rPr.isSetSz()) {
            ctrPr.setSz(rPr.getSz());
        }
        if (!ctrPr.isSetSzCs() && rPr.isSetSzCs()) {
            ctrPr.setSzCs(rPr.getSzCs());
        }
        if (!ctrPr.isSetLang() && rPr.isSetLang()) {
            ctrPr.setLang(rPr.getLang());
        }
    }

    /**
     * 给段落的 {@link CTPPr} 设置默认的段落样式
     *
     * @param ctpPr        段落的 CTPPr
     * @param ctpPrDefault 段落的默认样式
     */
    public void setParagraphDefaultStyles(CTPPr ctpPr, CTPPr ctpPrDefault) {
        // jc
        if (!ctpPr.isSetJc() && ctpPrDefault.isSetJc()) {
            ctpPr.setJc(ctpPrDefault.getJc());
        }
        // spacing
        if (!ctpPr.isSetSpacing() && ctpPrDefault.isSetSpacing()) {
            ctpPr.setSpacing(ctpPrDefault.getSpacing());
        }
        // shd
        if (!ctpPr.isSetShd() && ctpPrDefault.isSetShd()) {
            ctpPr.setShd(ctpPrDefault.getShd());
        }
    }
}
