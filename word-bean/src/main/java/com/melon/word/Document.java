package com.melon.word;

import com.melon.word.constants.Commons;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
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
     * 存放 XWPFDocument 和 Document 的 Map, 用于通过 XWPFDocument 获取 Document
     */
    private static final Map<XWPFDocument, Document> DOCUMENT_MAP = new HashMap<>(1);

    /**
     * @see org.apache.poi.xwpf.usermodel.XWPFDocument
     */
    private XWPFDocument xwpfDocument;

    /**
     * 存放 CTSectPr 的 List, sectPr 表示 section(部分), 有时候可以把 Word 文档看做是多个 seciton 拼接展示的.
     */
    private List<CTSectPr> sectPrList;

    /**
     * 使用 {@link java.io.InputStream} 创建一个 {@link Document} 对象
     *
     * @param inputStream {@link java.io.InputStream}
     * @return {@link Document}
     * @throws IOException 抛出{@link java.io.IOException}
     */
    public static Document generate(InputStream inputStream) throws IOException {
        if (inputStream == null) {
            throw new NullPointerException();
        }
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        Document document = new Document();
        document.xwpfDocument = xwpfDocument;
        // 放到 Map 中
        DOCUMENT_MAP.put(xwpfDocument, document);
        // 这里应该将 document 的 sectPr 放到 list 中
        if (xwpfDocument.getDocument().isSetBody()
                && xwpfDocument.getDocument().getBody().isSetSectPr()) {
            document.sectPrList.add(xwpfDocument.getDocument().getBody().getSectPr());
        }
        return document;
    }

    /**
     * 获取 {@link XWPFDocument} 对象
     *
     * @return 返回 XWPFDocument 对象
     */
    public XWPFDocument getDocument() {
        return xwpfDocument;
    }


    /**
     * 关闭资源对象
     *
     * @throws IOException 异常时抛出
     */
    public void close() throws IOException {
        this.xwpfDocument.close();
    }


    /**
     * 将 Map 中的数据替换到 {@link XWPFDocument} 对象中
     *
     * @param data key 可以是被替换的内容，比如配置 username，那么 key 就是 username；也可以配置 user.name, key 就是 user, 会自
     *             动将 user 对象中的 name 属性取出
     */
    public Document parse(Map<String, Object> data) {
        new Parser(xwpfDocument, data).parse();
        return this;
    }


    /**
     * 将资源保存到指定的 {@link OutputStream}
     *
     * @param os 指定的 OutputStream
     * @throws IOException 跑出 {@link IOException}
     */
    public void saveTo(OutputStream os) throws IOException {
        xwpfDocument.write(os);
    }


    /**
     * 判断是否存在配置
     *
     * @param text    内容
     * @param lenient 宽松的, true: 宽松的验证; false: 完全符合
     * @return true: 存在; false: 不存在
     */
    static boolean existsExpress(String text, boolean lenient) {
        Pattern pattern = SingletonPattern.pattern;
        Matcher matcher = pattern.matcher(text);
        if (lenient) {
            // 宽松的
            return matcher.find();
        } else {
            return matcher.matches();
        }
    }

    /**
     * 向 list 中加入此类型元素
     *
     * @param sectPr
     */
    public synchronized void addSectPr(CTSectPr sectPr) {
        if (this.sectPrList == null) {
            this.sectPrList = new ArrayList<>();
        }
        // 这里的 sectPr 应该是插入进去的, 因为 document 的 sectPr 应该永远在最后一个
        if (this.sectPrList.size() == 0) {
            this.sectPrList.add(sectPr);
        } else {
            this.sectPrList.add(this.sectPrList.size() - 1, sectPr);
        }
    }

    /**
     * 获取此类型元素的 List
     *
     * @return {@link List<CTSectPr>}
     */
    public List<CTSectPr> getSectPrList() {
        return this.sectPrList;
    }

    /**
     * 根据 {@link XWPFDocument} 获取 {@link Document} 对象
     *
     * @param document XWPFDocument 对象
     * @return Document 对象
     */
    public static Document getParentDocument(XWPFDocument document) {
        if (document == null) {
            throw new NullPointerException();
        }
        return DOCUMENT_MAP.get(document);
    }

    static class SingletonPattern {
        static Pattern pattern = Pattern.compile(Commons.DOLLAR_REGEX);
    }

}
