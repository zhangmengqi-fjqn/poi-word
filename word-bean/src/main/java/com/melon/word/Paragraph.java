package com.melon.word;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

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

    public Paragraph(XWPFParagraph paragraph) {
        this.paragraph = paragraph;
    }
}
