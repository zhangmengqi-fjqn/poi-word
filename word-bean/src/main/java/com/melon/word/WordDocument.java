package com.melon.word;

import com.melon.word.exceptions.WordException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;

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
     * 载入一个模板文件作为 Word
     *
     * @param inputStream {@link InputStream}
     */
    public WordDocument(InputStream inputStream) throws IOException, XmlException {
        document = new XWPFDocument(inputStream);
        setDefaultStyles(document);
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
     * @throws IOException when an error occurs
     * @see AutoCloseable
     */
    @Override
    public void close() throws IOException {
        // 实现 AutoCloseabe 接口的此方法是为了 try-with-resource 的语法糖准备的
        document.close();
    }
}
