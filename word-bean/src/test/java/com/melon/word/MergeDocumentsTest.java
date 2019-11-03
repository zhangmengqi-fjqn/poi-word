package com.melon.word;

import org.junit.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * @author zhaokai
 * @date 2019-11-03
 */
public class MergeDocumentsTest {

    @Test
    public void test() {
        final String word1 = Constant.TEST_PATH + "test1.docx";
        final String word2 = Constant.TEST_PATH + "test2.docx";
        try (
                WordDocument document = new WordDocument(word1);
                WordDocument document2 = new WordDocument(word2);
                OutputStream os = new FileOutputStream(Constant.TEST_PATH + "result.docx")
        ) {
            document.merge(document2);
            document.save(os);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
