package com.melon.word.util;

import com.melon.word.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author zhaokai
 * @date 2019-10-15
 */
public class CopyTableTest {

    static final String TMP_DIR = "/Users/zhaokai/Documents/test/";

    static final String TEST_PATH = TMP_DIR + "test.docx";

    @Test
    public void test() {
        try (FileInputStream fileInputStream = new FileInputStream(TEST_PATH);
             OutputStream os = new FileOutputStream(TMP_DIR + "result.docx");
        ) {
            Map<String, Object> data = new HashMap<>(16);
            data.put("user.name", "zhaokai");
            Document doc = Document.generate(fileInputStream);
            XWPFDocument document = doc.getDocument();
            doc.parse(data);

            // 旧的表格
            XWPFTable oldTable = document.getTables().get(0);

            // 创建新的 CTTbl ， table
            CTTbl ctTbl = CTTbl.Factory.newInstance();
            // 复制原来的CTTbl
            ctTbl.set(oldTable.getCTTbl());
            // 新增一个table，使用复制好的Cttbl
            XWPFTable newTable = new XWPFTable(ctTbl, document);
            document.setTable(0, newTable);


            doc.saveTo(os);
            System.out.println("successful!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
