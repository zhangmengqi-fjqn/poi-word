package com.melon.word;

import org.apache.commons.io.IOUtils;

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
        FileInputStream fileInputStream = null;
        OutputStream os = null;
        try {
            fileInputStream = new FileInputStream(path);
            Map<String, Object> data = new HashMap<>(16);
            data.put("user", new User("zhaokai", "男", 24));
            Document document = Document.generate(fileInputStream);
            document.compile(data);
            document.saveTo(os = new FileOutputStream(TMP_DIR + "result.docx"));
            System.out.println("successful!");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(fileInputStream, os);
        }
    }
}
