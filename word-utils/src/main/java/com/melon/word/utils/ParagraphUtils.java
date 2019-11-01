package com.melon.word.utils;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;

/**
 * 段落的工具类
 *
 * @author zhaokai
 * @date 2019-10-23
 * @see org.apache.poi.xwpf.usermodel.XWPFParagraph
 */
public class ParagraphUtils {

    private ParagraphUtils() {
    }

    /**
     * 拷贝 {@link CTPPr} 中的样式，将旧样式对象中的属性拷贝到新的样式对象中
     *
     * @param newCtPPr 新的样式对象
     * @param oldCtPPr 旧的样式对象
     */
    public static void setStyles(CTPPr newCtPPr, CTPPr oldCtPPr) {
        if (!newCtPPr.isSetPBdr()) {
            // pBdr
            newCtPPr.setPBdr(oldCtPPr.getPBdr());
        }
        if (!newCtPPr.isSetTabs()) {
            // tabs
            newCtPPr.setTabs(oldCtPPr.getTabs());
        }
        if (!newCtPPr.isSetSnapToGrid()) {
            // snapToGrid
            newCtPPr.setSnapToGrid(oldCtPPr.getSnapToGrid());
        }
        if (!newCtPPr.isSetSpacing()) {
            // spacing
            newCtPPr.setSpacing(oldCtPPr.getSpacing());
        }
        if (!newCtPPr.isSetJc()) {
            // jc
            newCtPPr.setJc(oldCtPPr.getJc());
        }
        if (oldCtPPr.isSetRPr()) {
            if (!newCtPPr.isSetRPr()) {
                newCtPPr.setRPr(oldCtPPr.getRPr());
            } else {
                if (newCtPPr.getRPr().isSetRStyle()) {
                    newCtPPr.getRPr().unsetRStyle();
                }
                setStyles(newCtPPr.getRPr(), oldCtPPr.getRPr());
            }
        }
    }

    public static void setStyles(CTParaRPr newCtPPr, CTParaRPr oldCtPPr) {

    }

}
