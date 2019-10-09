package com.melon.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;

/**
 * @author zhaokai
 * @since 2019-10-04
 */
public class Document {

    private Document() {
    }

    /**
     * @see org.apache.poi.xwpf.usermodel.XWPFDocument
     */
    private XWPFDocument xwpfDocument;

    public static Document compile(InputStream inputStream) throws IOException {
        if (inputStream == null) {
            throw new NullPointerException();
        }
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        Document document = new Document();
        document.xwpfDocument = xwpfDocument;
        return document;
    }

    public void close() throws IOException {
        this.xwpfDocument.close();
    }

}
