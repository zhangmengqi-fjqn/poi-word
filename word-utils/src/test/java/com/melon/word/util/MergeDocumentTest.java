package com.melon.word.util;

import com.melon.word.utils.DocumentUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import java.io.*;

/**
 * @author zhaokai
 * @date 2019-10-31
 */
public class MergeDocumentTest {

    @Test
    public void test() {
        final String testPath = com.melon.word.util.Test.TMP_DIR + "test1.docx";
        final String testPath1 = com.melon.word.util.Test.TMP_DIR + "test2.docx";
        try (InputStream is = new FileInputStream(testPath);
             InputStream is1 = new FileInputStream(testPath1);
             XWPFDocument document = new XWPFDocument(is);
             XWPFDocument document1 = new XWPFDocument(is1);
             OutputStream os = new FileOutputStream(com.melon.word.util.Test.TMP_DIR + "result.docx")) {
            DocumentUtils.merge(document, document1);
            document.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
