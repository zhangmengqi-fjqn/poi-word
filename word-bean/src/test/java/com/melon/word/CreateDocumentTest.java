package com.melon.word;

import org.junit.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * @author zhaokai
 * @date 2019-11-02
 */
public class CreateDocumentTest {

    @Test
    public void test() {
        final String resultPath = "/Users/zhaokai/Documents/test/result.docx";
        try (
                WordDocument document = new WordDocument();
                OutputStream os = new FileOutputStream(resultPath)
        ) {
            document.appendParagraph("这是第一个段落");
            document.appendParagraph("这是第二个段落");
            document.appendParagraph("这是第三个段落");
            document.appendParagraph("这是第四个段落");
            document.save(os);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
