package com.melon.word;

import org.junit.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * @author zhaokai
 * @date 2019-11-03
 */
public class MergeDocumentsTest {

    public static void main(String[] args) {
        final String word1 = Constant.TEST_PATH + "s.docx";
        final String word2 = Constant.TEST_PATH + "merged-common.docx";
        try (
                WordDocument document = new WordDocument(word1);
                WordDocument document2 = new WordDocument(word2);
                OutputStream os = new FileOutputStream(Constant.TEST_PATH + "result.docx")
        ) {
            document.merge(document2);
            document.save(os);
            System.out.println("successful!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
