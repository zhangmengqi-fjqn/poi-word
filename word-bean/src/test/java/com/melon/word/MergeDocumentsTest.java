package com.melon.word;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * @author zhaokai
 * @date 2019-11-03
 */
public class MergeDocumentsTest {

    public static void main(String[] args) {
        final String s = Constant.TEST_PATH + "s.docx";
        final String w = Constant.TEST_PATH + "w.docx";
        final String x = Constant.TEST_PATH + "x.docx";
        try (
                WordDocument ds = new WordDocument(s);
                WordDocument dw = new WordDocument(w);
                WordDocument dx = new WordDocument(x);
                OutputStream os = new FileOutputStream(Constant.TEST_PATH + "result.docx")
        ) {
            ds.merge(dw, true).merge(dx, true);
            ds.save(os);
            System.out.println("successful!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
