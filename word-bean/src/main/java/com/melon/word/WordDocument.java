package com.melon.word;

import com.google.common.collect.Lists;
import com.melon.word.extend.HeaderFooterPolicy;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.*;

/**
 * 适配器模式对 {@link org.apache.poi.xwpf.usermodel.XWPFDocument} 的封装
 *
 * @author zhaokai
 * @date 2019-11-01
 */
@Slf4j
public class WordDocument implements AutoCloseable {

    /**
     * @see XWPFDocument
     */
    private XWPFDocument document;

    /**
     * 在内部创建的输入流需要在内部关闭
     */
    private InputStream inputStream;

    private static final String HEADER = "header";
    private static final String FOOTER = "footer";

    /**
     * 载入一个模板文件作为 Word
     *
     * @param inputStream {@link InputStream}
     */
    public WordDocument(InputStream inputStream) throws IOException, XmlException {
        document = new XWPFDocument(inputStream);
        setDefaultStyles(document);
    }

    /**
     * 获取此对象
     *
     * @param document {@link XWPFDocument}
     */
    public WordDocument(XWPFDocument document) {
        if (document == null) {
            throw new NullPointerException();
        }
        this.document = document;
    }

    /**
     * 通过URL的方式创建文档对象
     *
     * @param url 文档的目录
     */
    public WordDocument(String url) throws IOException {
        if (!new File(url).exists()) {
            throw new FileNotFoundException();
        }
        inputStream = new FileInputStream(url);
        this.document = new XWPFDocument(inputStream);
    }

    /**
     * 创建一个空的 Word 文档
     */
    public WordDocument() throws IOException, XmlException {
        document = new XWPFDocument();
        setDefaultStyles(document);
    }

    /**
     * 给文档设置默认样式
     *
     * @param document {@link XWPFDocument}
     */
    private static void setDefaultStyles(XWPFDocument document) throws IOException, XmlException {
        CTStyles style;
        if (document.getStyles() == null) {
            // style 是空的，说明这个文档还没有创建样式的xml文件
            style = CTStyles.Factory.newInstance();
            XWPFStyles xwpfStyles = document.createStyles();
            xwpfStyles.setStyles(style);
        } else {
            style = document.getStyle();
        }
        // 新建或者直接获取已有的 CTDocDefaults
        CTDocDefaults ctDocDefaults = style.isSetDocDefaults() ? style.getDocDefaults() : style.addNewDocDefaults();
        // 新建或者直接获取已有的 CTRPrDefault
        CTRPrDefault ctrPrDefault = ctDocDefaults.isSetRPrDefault() ? ctDocDefaults.getRPrDefault() : ctDocDefaults.addNewRPrDefault();
        // 新建或者直接获取已有的 CTRPr
        CTRPr ctrPr = ctrPrDefault.isSetRPr() ? ctrPrDefault.getRPr() : ctrPrDefault.addNewRPr();
        // 获取并设置字体
        CTFonts ctFonts = ctrPr.isSetRFonts() ? ctrPr.getRFonts() : ctrPr.addNewRFonts();
        ctFonts.setAscii("宋体");
        ctFonts.setEastAsia("宋体");
        // 设置字体大小
        CTHpsMeasure ctHpsMeasure = ctrPr.isSetSz() ? ctrPr.getSz() : ctrPr.addNewSz();
        ctHpsMeasure.setVal(new BigInteger("24"));
    }

    /**
     * 向文档中插入一个新的段落
     *
     * @param paragraphText 段落的内容
     * @return {@link XWPFParagraph}
     */
    public Paragraph appendParagraph(String paragraphText) {
        XWPFParagraph paragraph = this.document.createParagraph();
        if (paragraphText != null) {
            // 段落内容不为空，就把内容放入段落
            paragraph.createRun().setText(paragraphText, 0);
        }
        return new Paragraph(paragraph);
    }

    /**
     * 将文档保存到输出流中
     *
     * @param outputStream 输出流
     * @throws IOException when an error occurs
     */
    public void save(OutputStream outputStream) throws IOException {
        document.write(outputStream);
    }

    /**
     * 调用 close 方法关闭 XWPFDocument 对象，可以 使用 try-with-source 语法
     *
     * @throws IOException when an error occurs
     * @see AutoCloseable
     */
    @Override
    public void close() throws IOException {
        // 实现 AutoCloseabe 接口的此方法是为了 try-with-resource 的语法糖准备的
        document.close();
        if (inputStream != null) {
            inputStream.close();
        }
    }

    /**
     * 合并文档
     *
     * @param wordDocument 被合并的文档
     */
    public WordDocument merge(WordDocument wordDocument) throws IOException, XmlException {
        // 复制之前将 sectPr 放到此文档最后一个段落
        if (this.document.getDocument().getBody().isSetSectPr()) {
            XWPFParagraph paragraph = this.document.createParagraph();
            paragraph.getCTP().addNewPPr().setSectPr(this.document.getDocument().getBody().getSectPr());
        }

        // 将 wordDocument 中的元素全部复制到 this 对象中
        Iterator<IBodyElement> bodyElementsIterator = wordDocument.document.getBodyElementsIterator();
        while (bodyElementsIterator.hasNext()) {
            // 每个元素
            IBodyElement bodyElement = bodyElementsIterator.next();
            // 获取类型并判断
            BodyElementType elementType = bodyElement.getElementType();
            if (elementType == BodyElementType.TABLE) {
                // 合并表格
                copyTableToThis((XWPFTable) bodyElement);
            } else if (elementType == BodyElementType.PARAGRAPH) {
                // 合并段落
                copyParagraphToThis((XWPFParagraph) bodyElement, wordDocument.document);
            }
        }

        // 将被合并的文档的 sectPr 放入基础文档中
        if (wordDocument.document.getDocument().getBody().isSetSectPr()) {
            this.document.getDocument().getBody().setSectPr(wordDocument.document.getDocument().getBody().getSectPr());

            // 设置原来的页眉和页脚
            copyReferencesToThisDocument(wordDocument.document);
        }
        return this;
    }

    /**
     * 将 document 的页眉页脚等复制到基础文档中
     *
     * @param document 被合并的文档
     */
    private void copyReferencesToThisDocument(XWPFDocument document) {
        // 获取 main 文档的 sectPr
        CTSectPr sectPr = this.document.getDocument().getBody().getSectPr();
        // 获取 reference 的信息
        // 总的 reference 的 map
        Map<String, List<List<XWPFParagraph>>> refMap = new HashMap<>(16);
        // 页眉的
        List<CTHdrFtrRef> headerReferenceList = sectPr.getHeaderReferenceList();
        if (CollectionUtils.isNotEmpty(headerReferenceList)) {
            for (CTHdrFtrRef ctHdrFtrRef : headerReferenceList) {
                POIXMLDocumentPart documentPart = document.getRelationById(ctHdrFtrRef.getId());
                List<XWPFParagraph> paragraphs = ((XWPFHeader) documentPart).getParagraphs();
                if (refMap.containsKey(HEADER)) {
                    // 已经存在，追加
                    refMap.get(HEADER).add(paragraphs);
                } else {
                    List<List<XWPFParagraph>> paragraphLists = new ArrayList<>();
                    paragraphLists.add(paragraphs);
                    refMap.put(HEADER, paragraphLists);
                }
            }
        } else {
            // 没有就加一个空的段落
            List<List<XWPFParagraph>> paragraphLists = new ArrayList<>();
            paragraphLists.add(Lists.newArrayList(new Paragraph(this.document).getParagraph()));
            refMap.put(HEADER, paragraphLists);
        }
        // 页脚的
        List<CTHdrFtrRef> footerReferenceList = sectPr.getFooterReferenceList();
        if (CollectionUtils.isNotEmpty(footerReferenceList)) {
            for (CTHdrFtrRef ctHdrFtrRef : footerReferenceList) {
                POIXMLDocumentPart documentPart = document.getRelationById(ctHdrFtrRef.getId());
                List<XWPFParagraph> paragraphs = ((XWPFFooter) documentPart).getParagraphs();
                if (refMap.containsKey(FOOTER)) {
                    // 已经存在，追加
                    refMap.get(FOOTER).add(paragraphs);
                } else {
                    List<List<XWPFParagraph>> paragraphLists = new ArrayList<>();
                    paragraphLists.add(paragraphs);
                    refMap.put(FOOTER, paragraphLists);
                }
            }
        } else {
            // 没有就加一个空的段落
            List<List<XWPFParagraph>> paragraphLists = new ArrayList<>();
            paragraphLists.add(Lists.newArrayList(new Paragraph(this.document).getParagraph()));
            refMap.put(FOOTER, paragraphLists);
        }

        // 获取完之后就把 main 中的 reference 全部移除
        if (CollectionUtils.isNotEmpty(headerReferenceList)) {
            for (int i = 0; i < headerReferenceList.size(); i++) {
                sectPr.removeHeaderReference(i);
            }
        }
        if (CollectionUtils.isNotEmpty(footerReferenceList)) {
            for (int i = 0; i < footerReferenceList.size(); i++) {
                sectPr.removeFooterReference(i);
            }
        }

        // 总的, 不为空就取复制
        if (MapUtils.isNotEmpty(refMap)) {
            // 页眉信息不为空, 就去复制
            for (Map.Entry<String, List<List<XWPFParagraph>>> entry : refMap.entrySet()) {
                String key = entry.getKey();
                List<List<XWPFParagraph>> paragraphLists = entry.getValue();
                if (Objects.equals(HEADER, key)) {
                    // 页眉
                    for (List<XWPFParagraph> paragraphList : paragraphLists) {
                        addHeader(this.document, CollectionUtils.isEmpty(paragraphList) ?
                                Lists.newArrayList(new Paragraph(this.document).getParagraph()) : paragraphList);
                    }
                } else {
                    // 页脚
                    for (List<XWPFParagraph> paragraphList : paragraphLists) {
                        addFooter(this.document, CollectionUtils.isEmpty(paragraphList) ?
                                Lists.newArrayList(new Paragraph(this.document).getParagraph()) : paragraphList);
                    }
                }
            }
        }
    }

    /**
     * 将段落复制到 this 中
     *
     * @param paragraph    段落
     * @param xwpfDocument 被合并的文档
     */
    private void copyParagraphToThis(XWPFParagraph paragraph, XWPFDocument xwpfDocument) throws IOException, XmlException {
        CTP subCtp = paragraph.getCTP();
        if (subCtp.isSetPPr() && subCtp.getPPr().isSetSectPr()) {
            // 如果这个段落带了 sectPr 则不用处理
            return;
        }
        // 使用 xmlObject 创建一个 paragraph 段落
        XWPFParagraph newPara = new XWPFParagraph(subCtp, this.document);
        // 在 mainDocument 创建一个空的段落
        XWPFParagraph newParagraph = this.document.createParagraph();
        // 获取新的段落在元素中的位置
        int elementPosition = this.document.getPosOfParagraph(newParagraph);
        // 获取新的段落在段落list中的位置
        int paragraphPosition = this.document.getParagraphPos(elementPosition);
        // 设置默认样式
        new Paragraph(newPara).setDocumentDefaultStyles(xwpfDocument.getStyle());
        // 将使用 xmlObject 创建的段落 set 到 mainDocument 创建的空段落上
        this.document.setParagraph(newPara, paragraphPosition);
    }

    /**
     * 将表格在 this 中创建一份
     *
     * @param table 表格
     */
    private void copyTableToThis(XWPFTable table) {
        CTTbl subCtTbl = table.getCTTbl();
        // 使用 xmlObject 创建一个 table 表格
        XWPFTable newCreatedTable = new XWPFTable(subCtTbl, this.document);
        // 在 mainDocument 创建一个空的表格
        XWPFTable newTable = this.document.createTable();
        // 获取新的表格在元素中的位置
        int elementPosition = this.document.getPosOfTable(newTable);
        // 获取新的表格在表格list中的位置
        int tablePosition = this.document.getTablePos(elementPosition);
        // 将使用 xmlObject 创建的表格 set 到 mainDocument 创建的空表格上
        this.document.setTable(tablePosition, newCreatedTable);
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

}
