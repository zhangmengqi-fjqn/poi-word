package com.melon.word.util;

import com.melon.word.Document;
import com.melon.word.utils.DocumentUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author zhaokai
 * @date 2019-10-23
 */
public class AddHeaderToDocument {

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

            // addHeader(doc);

            addParagraph(doc);


            doc.saveTo(os);
            System.out.println("successful!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void addParagraph(Document doc) {
        XWPFDocument document = doc.getDocument();
        // addHeader("第一个页眉", document.getDocument().getBody().getSectPr(), document);
        addPText(document, "这是段落1");
        DocumentUtils.insertNextPageChar(document);
        addPText(document, "这是段落2");
        addHeader("第2个页眉", Document.getParentDocument(document).getSectPrList().get(0), document);
        DocumentUtils.insertNextPageChar(document);
        addPText(document, "这是段落3");
        DocumentUtils.insertNextPageChar(document);
        addHeader("第3个页眉", Document.getParentDocument(document).getSectPrList().get(2), document);
        addPText(document, "这是段落4");
    }

    private void addHeader(String msg, CTSectPr sectPr, XWPFDocument document) {
        CTP ctp = CTP.Factory.newInstance();
        CTR ctr = ctp.addNewR();
        CTText ctText = ctr.addNewT();
        ctText.setStringValue(msg);
        XWPFParagraph paragraph = new XWPFParagraph(ctp, document);
        List<XWPFParagraph> paragraphList = new ArrayList<>(1);
        paragraphList.add(paragraph);
        DocumentUtils.addHeader(document, sectPr, paragraphList);
    }

    private void addPText(XWPFDocument document, String msg) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(msg, 0);
    }

}
