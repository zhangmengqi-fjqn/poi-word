package com.melon.word.util;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;

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
public class TestSignature {

    private static final String TMP_DIR = "D:\\work\\temp\\tl\\";

    public static void main(String[] args) {
        try (FileInputStream fileInputStream = new FileInputStream(TMP_DIR + "test.docx");
             OutputStream os = new FileOutputStream(TMP_DIR + "result.docx");
        ) {
            Map<String, Object> data = new HashMap<>(16);
            data.put("makerSign", new PictureRenderData(120, 60, TMP_DIR + "1.png"));
// 依旧损坏
//            Configure.ConfigureBuilder builder = Configure.newBuilder();
//            builder.addPlugin('%', new PictureRenderPolicyCustomer());
//            Configure configure = builder.build();
//            XWPFTemplate template = XWPFTemplate.compile(fileInputStream, configure).render(data);
            XWPFTemplate template = XWPFTemplate.compile(fileInputStream).render(data);
            template.write(os);
            os.flush();
            os.close();
            template.close();
            System.out.println("successful!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
