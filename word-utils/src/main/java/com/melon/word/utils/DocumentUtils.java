package com.melon.word.utils;

import com.melon.word.Document;
import com.melon.word.extend.HeaderFooterPolicy;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

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
     * 向 sectPr 中插入个页脚
     *
     * @param document   {@link XWPFDocument}
     * @param sectPr     {@link CTSectPr}
     * @param paragraphs {@link List<XWPFParagraph>}
     */
    public static void addFooter(XWPFDocument document, CTSectPr sectPr, List<XWPFParagraph> paragraphs) {
        HeaderFooterPolicy policy = new HeaderFooterPolicy(document, sectPr);
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, CollectionUtils.isEmpty(paragraphs) ? new XWPFParagraph[]{} : paragraphs.toArray(new XWPFParagraph[]{}));
    }


    /**
     * 向文档中插入个页脚
     *
     * @param document   {@link XWPFDocument}
     * @param paragraphs {@link List<XWPFParagraph>}
     */
    public static void addFooter(XWPFDocument document, List<XWPFParagraph> paragraphs) {
        // sectPr 为 null 时, 将会自动获取 document 的 sectPr
        addFooter(document, null, paragraphs);
    }

    /**
     * 合并文档
     *
     * @param mainDocument 主要文档
     * @param subDocument  下一个文档
     */
    public static void merge(XWPFDocument mainDocument, XWPFDocument subDocument) {
        // 获取第二个文档的正文内容的元素
        List<IBodyElement> bodyElements = subDocument.getBodyElements();
        for (IBodyElement bodyElement : bodyElements) {
            BodyElementType elementType = bodyElement.getElementType();
            if (elementType == BodyElementType.PARAGRAPH) {
                // 处理段落
                XWPFParagraph subParagraph = (XWPFParagraph) bodyElement;
                CTP subCtp = subParagraph.getCTP();
                // 在 mainDocument 创建一个空的段落
                XWPFParagraph newParagraph = mainDocument.createParagraph();
                if (subCtp.isSetPPr() && subCtp.getPPr().isSetSectPr()) {
                    // 处理 sectPr
                    CTPPr ctpPr = newParagraph.getCTP().isSetPPr() ? newParagraph.getCTP().getPPr() : newParagraph.getCTP().addNewPPr();
                    CTSectPr oldSectPr = subCtp.getPPr().getSectPr();
                    ctpPr.setSectPr(oldSectPr);
                    // 先清空页眉页脚
                    CTSectPr sectPr = ctpPr.getSectPr();
                    int headerSize = sectPr.sizeOfHeaderReferenceArray();
                    for (int i = 0; i < headerSize; i++) {
                        sectPr.removeHeaderReference(i);
                    }
                    int footerSize = sectPr.sizeOfFooterReferenceArray();
                    for (int i = 0; i < footerSize; i++) {
                        sectPr.removeFooterReference(i);
                    }
                    XWPFStyles styles = subDocument.getStyles();
                    // 页眉
                    List<CTHdrFtrRef> headerReferenceList = oldSectPr.getHeaderReferenceList();
                    XWPFStyle headerStyle = styles.getStyleWithName("header");
                    CTPPr headerPPr = headerStyle.getCTStyle().getPPr();
                    for (CTHdrFtrRef ctHdrFtrRef : headerReferenceList) {
                        POIXMLDocumentPart oldDocumentPart = subDocument.getRelationById(ctHdrFtrRef.getId());
                        if (oldDocumentPart instanceof XWPFHeader) {
                            List<XWPFParagraph> paragraphs = ((XWPFHeader) oldDocumentPart).getParagraphs();
                            for (XWPFParagraph paragraph : paragraphs) {
                                if (!paragraph.getCTP().isSetPPr()) {
                                    paragraph.getCTP().setPPr(headerPPr);
                                } else {
                                    CTPPr pPr = paragraph.getCTP().getPPr();
                                    if (pPr.getPStyle() != null) {
                                        pPr.unsetPStyle();
                                        ParagraphUtils.setStyles(pPr, headerPPr);
                                    }
                                }
                            }
                            addHeader(mainDocument, sectPr, paragraphs);
                        }
                    }
                    // 页脚
                    XWPFStyle footerStyle = styles.getStyleWithName("footer");
                    CTPPr footerPPr = footerStyle.getCTStyle().getPPr();
                    List<CTHdrFtrRef> footerReferenceList = oldSectPr.getFooterReferenceList();
                    for (CTHdrFtrRef ctHdrFtrRef : footerReferenceList) {
                        POIXMLDocumentPart oldDocumentPart = subDocument.getRelationById(ctHdrFtrRef.getId());
                        if (oldDocumentPart instanceof XWPFFooter) {
                            List<XWPFParagraph> paragraphs = ((XWPFFooter) oldDocumentPart).getParagraphs();
                            for (XWPFParagraph paragraph : paragraphs) {
                                if (!paragraph.getCTP().isSetPPr()) {
                                    paragraph.getCTP().setPPr(footerPPr);
                                } else {
                                    CTPPr pPr = paragraph.getCTP().getPPr();
                                    if (pPr.getPStyle() != null) {
                                        pPr.unsetPStyle();
                                        ParagraphUtils.setStyles(pPr, footerPPr);
                                    }
                                }
                            }
                            addFooter(mainDocument, sectPr, ((XWPFFooter) oldDocumentPart).getParagraphs());
                        }
                    }
                    continue;
                }
                // 使用 xmlObject 创建一个 paragraph 段落
                XWPFParagraph paragraph = new XWPFParagraph(subCtp, mainDocument);
                // 获取新的段落在元素中的位置
                int elementPosition = mainDocument.getPosOfParagraph(newParagraph);
                // 获取新的段落在段落list中的位置
                int paragraphPosition = mainDocument.getParagraphPos(elementPosition);
                // 将使用 xmlObject 创建的段落 set 到 mainDocument 创建的空段落上
                mainDocument.setParagraph(paragraph, paragraphPosition);
            } else if (elementType == BodyElementType.TABLE) {
                // 处理表格
                XWPFTable subTable = (XWPFTable) bodyElement;
                CTTbl subCtTbl = subTable.getCTTbl();
                // 使用 xmlObject 创建一个 table 表格
                XWPFTable table = new XWPFTable(subCtTbl, mainDocument);
                // 在 mainDocument 创建一个空的表格
                XWPFTable newTable = mainDocument.createTable();
                // 获取新的表格在元素中的位置
                int elementPosition = mainDocument.getPosOfTable(newTable);
                // 获取新的表格在表格list中的位置
                int tablePosition = mainDocument.getTablePos(elementPosition);
                // 将使用 xmlObject 创建的表格 set 到 mainDocument 创建的空表格上
                mainDocument.setTable(tablePosition, table);
            }
        }


        System.out.println("ok");
    }

}
