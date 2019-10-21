package com.melon.word;

import com.melon.word.constants.Commons;
import org.apache.commons.jexl3.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * @author zhaokai
 * @date 2019-10-16
 */
public class Parser {

    private static final String UNDEFINED_VARIABLE = "undefined variable";

    /**
     * Jexl 引擎
     */
    private JexlEngine jexlEngine;

    /**
     * 数据的 context
     */
    private MapContext mapContext;

    /**
     * 带解析的段落 {@link XWPFParagraph}
     */
    private List<XWPFParagraph> paragraphs;

    /**
     * 类似于缓存的 Map
     */
    private Map<String, String> cacheMap;

    public Parser(XWPFDocument document, Map<String, Object> data) {
        this.mapContext = new MapContext(data);
        // 初始化
        this.jexlEngine = SingleJexlEngine.jexlEngine;
        cacheMap = new HashMap<>(16);
        // 查询待解析的段落
        paragraphs = getParsingParagraph(document.getParagraphs());
    }

    /**
     * 开始解析段落
     */
    public void parse() {
        // 解析段落
        parseParagraph();
    }

    private void parseParagraph() {
        AtomicBoolean whole = new AtomicBoolean(true);
        // 全局 XWPFRun
        XWPFRun run = null;
        // 全局的表达式
        StringBuffer express = new StringBuffer();
        // 一个 XWPFRun 中的内容: 'mynameis${user.' 类似的, 则可以保留 'mynaneis'
        String prefix = "";
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            int runSize = runs.size();
            for (int i = 0; i < runSize; i++) {
                XWPFRun tmpRun = runs.get(i);
                String runText = tmpRun.text();
                if (Document.existsExpress(runText, false)) {
                    runText = runText.substring(2, runText.length() - 1);
                    tmpRun.setText(getExpressionValue(runText), 0);
                } else {
                    int beginIndex = runText.indexOf(Commons.LEFT_BRACKETS) + 2;
                    int endIndex = runText.indexOf(Commons.RIGHT_BRACKETS);
                    if (beginIndex != 1 && endIndex != -1 && beginIndex < endIndex) {
                        // 此 run 包含表达式
                        String exp = runText.substring(beginIndex, endIndex);
                        tmpRun.setText(runText.replace(Commons.LEFT_BRACKETS + exp + Commons.RIGHT_BRACKETS, getExpressionValue(exp)), 0);
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
                            if (run == null) {
                                run = tmpRun;
                            }
                            run.setText(prefix + getExpressionValue(express.toString()), 0);
                            // 清空内容
                            express.delete(0, express.length());
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

    /**
     * 根据表达式获取值
     *
     * @param expression 表达式
     * @return 返回表达式对应的值
     */
    private String getExpressionValue(String expression) {
        String value = cacheMap.get(expression);
        if (value != null) {
            return value;
        }
        JexlExpression jexlExpression = jexlEngine.createExpression(expression);
        Object evaluate = null;
        try {
            evaluate = jexlExpression.evaluate(mapContext);
        } catch (JexlException e) {
            if (e.getMessage().contains(UNDEFINED_VARIABLE)) {
                return Commons.LEFT_BRACKETS + expression + Commons.RIGHT_BRACKETS;
            }
        }
        // 判断一下 evaluate
        value = evaluate == null ? "" : evaluate.toString();
        // 放入 Map
        cacheMap.put(expression, value);
        return value;
    }

    /**
     * 搜索可以解析的 XWPFParagraph
     *
     * @param paragraphs 所有的段落对象
     * @return 待解析的 List
     */
    private static List<XWPFParagraph> getParsingParagraph(List<XWPFParagraph> paragraphs) {
        // 待解析的 List
        List<XWPFParagraph> waitParseList = new ArrayList<>(paragraphs.size());
        for (XWPFParagraph paragraph : paragraphs) {
            if (!Document.existsExpress(paragraph.getText(), true)) {
                continue;
            }
            waitParseList.add(paragraph);
        }
        return waitParseList;
    }

    /**
     * 单例模式获取 {@link JexlEngine} 对象
     */
    private static class SingleJexlEngine {
        public static JexlEngine jexlEngine = new JexlBuilder().create();
    }
}
