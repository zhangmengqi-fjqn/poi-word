package com.melon.word;

import com.melon.word.exceptions.WordException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;

/**
 * 适配器模式对 {@link org.apache.poi.xwpf.usermodel.XWPFDocument} 的封装
 *
 * @author zhaokai
 * @date 2019-11-01
 */
@Slf4j
public class WordDocument {

    /**
     * @see XWPFDocument
     */
    private XWPFDocument document;

    /**
     * 载入一个模板文件作为 Word
     *
     * @param inputStream {@link InputStream}
     */
    public WordDocument(InputStream inputStream) {
        try {
            document = new XWPFDocument(inputStream);
        } catch (IOException e) {
            log.error("Error creating WordDocument.", e);
            throw new WordException("Error creating WordDocument.", e);
        }
    }

    /**
     * 创建一个空的 Word 文档
     */
    public WordDocument() {
        document = new XWPFDocument();
    }
}
