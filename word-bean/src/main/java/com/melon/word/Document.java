package com.melon.word;

import com.melon.word.constants.Commons;
import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlExpression;
import org.apache.commons.jexl3.MapContext;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 文档
 *
 * @author zhaokai
 * @since 2019-10-04
 */
public class Document {

    private Document() {
    }

    /**
     * @see org.apache.poi.xwpf.usermodel.XWPFDocument
     */
    private XWPFDocument xwpfDocument;

    /**
     * 使用 {@link java.io.InputStream} 创建一个 {@link Document} 对象
     *
     * @param inputStream {@link java.io.InputStream}
     * @return {@link Document}
     * @throws IOException {@link java.io.IOException}
     */
    public static Document generate(InputStream inputStream) throws IOException {
        if (inputStream == null) {
            throw new NullPointerException();
        }
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        Document document = new Document();
        document.xwpfDocument = xwpfDocument;
        return document;
    }

    public void close() throws IOException {
        this.xwpfDocument.close();
    }


    /**
     * 将 Map 中的数据替换到 {@link XWPFDocument} 对象中
     *
     * @param data key 可以是被替换的内容，比如配置 username，那么 key 就是 username；也可以配置 user.name, key 就是 user, 会自
     *             动将 user 对象中的 name 属性取出
     */
    public void compile(Map<String, Object> data) {
        JexlEngine jexlEngine = new JexlBuilder().create();
        MapContext mapContext = new MapContext(data);
        data.forEach((key, object) -> {
            JexlExpression expression = jexlEngine.createExpression(key);
            Object evaluate = expression.evaluate(mapContext);
            System.out.println("evaluate = " + evaluate);
        });
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        AtomicBoolean whole = new AtomicBoolean(true);
        // 全局 XWPFRun
        XWPFRun run = null;
        // 全局的表达式
        StringBuffer express = new StringBuffer();
        // 一个 XWPFRun 中的内容: 'mynameis${user.' 类似的, 则可以保留 'mybaneis'
        String prefix = "";
        for (XWPFParagraph paragraph : paragraphs) {
            String paragraphText = paragraph.getText();
            if (!existsExpress(paragraphText)) {
                continue;
            }
            List<XWPFRun> runs = paragraph.getRuns();
            int runSize = runs.size();
            for (int i = 0; i < runSize; i++) {
                XWPFRun tmpRun = runs.get(i);
                String runText = tmpRun.text();
                if (existsExpress(runText)) {
                    runText = runText.substring(2, runText.length() - 1);
                    JexlExpression expression = jexlEngine.createExpression(runText);
                    Object evaluate = expression.evaluate(mapContext);
                    tmpRun.setText(evaluate.toString(), 0);
                } else {
                    int beginIndex = runText.indexOf("${") + 2;
                    int endIndex = runText.indexOf("}");
                    if (beginIndex != 1 && endIndex != -1 && beginIndex >= endIndex) {
                        // 此 run 包含表达式
                        String exp = runText.substring(beginIndex, endIndex);
                        JexlExpression expression = jexlEngine.createExpression(exp);
                        Object evaluate = expression.evaluate(mapContext);
                        tmpRun.setText(evaluate.toString(), 0);
                    } else if (beginIndex != 1) {
                        // 存在 ${
                        express.append(runText.substring(beginIndex));
                        if (beginIndex > 2) {
                            prefix = runText.substring(0, beginIndex - 2);
                        }
                        run = tmpRun;
                        whole.set(false);
                    } else if (endIndex != -1) {
                        if (!whole.get()) {
                            express.append(runText, 0, endIndex);
                            JexlExpression expression = jexlEngine.createExpression(express.toString());
                            // 清空内容
                            express.delete(0, express.length());
                            Object evaluate = expression.evaluate(mapContext);
                            if (run == null) {
                                run = tmpRun;
                            }
                            run.setText(prefix + evaluate.toString(), 0);
                            whole.set(true);
                            // 置空
                            prefix = "";
                            tmpRun.setText(runText.substring(endIndex + 1), 0);
                        }
                    } else {
                        if (!whole.get()) {
                            express.append(runText);
                            paragraph.removeRun(i--);
                            runSize--;
                        }
                    }
                }
            }
        }
    }

    public void saveTo(OutputStream os) throws IOException {
        xwpfDocument.write(os);
    }

    private boolean existsExpress(String text) {
        Pattern pattern = Pattern.compile(Commons.DOLLAR_REGEX);
        Matcher matcher = pattern.matcher(text);
        return matcher.find();
    }

}
