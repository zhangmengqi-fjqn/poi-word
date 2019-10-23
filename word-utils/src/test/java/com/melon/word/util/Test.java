package com.melon.word.util;

import com.melon.word.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
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
public class Test {

    static final String TMP_DIR = "/Users/zhaokai/Documents/test/";

    static final String TEST_PATH = TMP_DIR + "test.docx";

    public static void main(String[] args) {
        final String path = TMP_DIR + "test.docx";
        try (FileInputStream fileInputStream = new FileInputStream(path);
             OutputStream os = new FileOutputStream(TMP_DIR + "result.docx");
        ) {
            Map<String, Object> data = new HashMap<>(16);
            data.put("user.name", "zhaokai");
            Document document = Document.generate(fileInputStream);
            XWPFDocument document1 = document.getDocument();
            document.parse(data);


            XWPFTable table = document1.getTables().get(0);
            CTTbl ctTbl = table.getCTTbl();
            XWPFTable newTable = document1.createTable();
            newTable.getCTTbl().setTblGrid(ctTbl.getTblGrid());
            newTable.getCTTbl().setTblPr(ctTbl.getTblPr());
            newTable.getCTTbl().setTrArray(ctTbl.getTrArray());

            document.saveTo(os);
            System.out.println("successful!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
