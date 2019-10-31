package com.melon.word.utils;

import com.melon.word.Document;
import com.melon.word.extend.HeaderFooterPolicy;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.util.List;

/**
 * 文档的工具类
 *
 * @author zhaokai
 * @date 2019-10-22
 * @see XWPFDocument
 */
public class DocumentUtils {

    private DocumentUtils() {
        // 私有化构造方法, 因为这是一个工具类
    }

    /**
     * 给 {@link XWPFDocument} 设置下一页
     * <br />
     * 也只有设置下一页形式的分页符才可以分隔不同页面的表头
     *
     * @param document {@link XWPFDocument} 对象
     */
    public static void insertNextPageChar(XWPFDocument document) {
        // 首先获取 document 的 Section 信息
        CTBody body = document.getDocument().getBody();
        // 放心, 这个 body 肯定不为空, 否则这个文档就有问题了
        XWPFParagraph paragraph = document.createParagraph();
        // 新创建的段落肯定没有 PPr, 所以需要新创建一个
        CTPPr ctpPr = paragraph.getCTP().addNewPPr();
        // 这一句其实就是设置下一页的分页符了
        CTSectPr sectPr = ctpPr.addNewSectPr();
        // 先加入到 document 的 List 中
        Document parent = Document.getParentDocument(document);
        if (parent != null) {
            parent.addSectPr(sectPr);
        }
        if (!body.isSetSectPr()) {
            // 文档没设置了 sectPr
            return;
        }
        // 把文档上的 sectPr 中的某些属性赋值给新创建的段落的 sectPr
        CTSectPr bodySectPr = body.getSectPr();
        sectPr.setPgSz(bodySectPr.getPgSz());
        sectPr.setPgMar(bodySectPr.getPgMar());
        sectPr.setCols(bodySectPr.getCols());
        sectPr.setDocGrid(bodySectPr.getDocGrid());
    }


    /**
     * 向 sectPr 中插入个页眉
     *
     * @param document   {@link XWPFDocument}
     * @param sectPr     {@link CTSectPr}
     * @param paragraphs {@link List<XWPFParagraph>}
     */
    public static void addHeader(XWPFDocument document, CTSectPr sectPr, List<XWPFParagraph> paragraphs) {
        HeaderFooterPolicy policy = new HeaderFooterPolicy(document, sectPr);
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, CollectionUtils.isEmpty(paragraphs) ? new XWPFParagraph[]{} : paragraphs.toArray(new XWPFParagraph[]{}));
    }


    /**
     * 向文档中插入个页眉
     *
     * @param document   {@link XWPFDocument}
     * @param paragraphs {@link List<XWPFParagraph>}
     */
    public static void addHeader(XWPFDocument document, List<XWPFParagraph> paragraphs) {
        // sectPr 为 null 时, 将会自动获取 document 的 sectPr
        addHeader(document, null, paragraphs);
    }

    /**
     * 合并文档
     *
     * @param mainDocument 主要文档
     * @param subDocument  下一个文档
     */
    public static void merge(XWPFDocument mainDocument, XWPFDocument subDocument) {

    }

}
