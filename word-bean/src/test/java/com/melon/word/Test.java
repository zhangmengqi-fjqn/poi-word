package com.melon.word;

import com.melon.word.utils.DocumentUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
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

    private static final String TMP_DIR = "/Users/zhaokai/Documents/test/";

    public static void main(String[] args) {
        final String path = TMP_DIR + "test.docx";
        try (FileInputStream fileInputStream = new FileInputStream(path);
             OutputStream os = new FileOutputStream(TMP_DIR + "result.docx");
        ) {
            Map<String, Object> data = new HashMap<>(16);
            data.put("user", new User("zhaokai", "男", 24));
            Document document = Document.generate(fileInputStream);
            XWPFDocument document1 = document.getDocument();
            document.parse(data);


            XWPFTable table = document1.getTables().get(0);
            CTTbl ctTbl = table.getCTTbl();
            XWPFTable newTable = document1.createTable();
            newTable.getCTTbl().setTblGrid(ctTbl.getTblGrid());
            newTable.getCTTbl().setTblPr(ctTbl.getTblPr());
            newTable.getCTTbl().setTrArray(ctTbl.getTrArray());

            newTable.getRows().get(0).getTableCells().get(0).setText("我就是那个内容");


            document.saveTo(os);
            System.out.println("successful!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
