package com.melon.word.util;

import com.melon.word.Document;
import com.melon.word.utils.TableUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author zhaokai
 * @date 2019-10-23
 */
public class InsertNewTable {

    @Test
    public void test() {
        try (FileInputStream fileInputStream = new FileInputStream(com.melon.word.util.Test.TEST_PATH);
             OutputStream os = new FileOutputStream(com.melon.word.util.Test.TMP_DIR + "result.docx");
        ) {
            Map<String, Object> data = new HashMap<>(16);
            data.put("user.name", "zhaokai");
            Document doc = Document.generate(fileInputStream);
            XWPFDocument document = doc.getDocument();
            doc.parse(data);

            TableUtils.appendTable(document, 3, 2);

            doc.saveTo(os);
            System.out.println("successful!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
