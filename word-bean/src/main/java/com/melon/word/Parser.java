package com.melon.word;

import com.melon.word.constants.Commons;
import org.apache.commons.jexl3.*;
import org.apache.poi.xwpf.usermodel.*;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;

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
     * 待解析的段落 {@link XWPFParagraph}
     */
    private List<XWPFParagraph> paragraphs;

    /**
     * 待解析的表格 {@link XWPFTable}
     */
    private List<XWPFTable> tables;

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
        paragraphs = getParsingParagraphs(document.getParagraphs());
        // 查询待解析的表格
        tables = getParsingTables(document.getTables());
    }

    /**
     * 开始解析段落
     */
    public void parse() {
        // 解析段落
        parseParagraphs(this.paragraphs);
        // 解析表格
        parseTables(this.tables);
    }

    /**
     * 解析表格
     */
    private void parseTables(List<XWPFTable> tables) {
        for (XWPFTable table : tables) {
            // 这个表格的行
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell tableCell : tableCells) {
                    parseParagraphs(tableCell.getParagraphs());
                }
            }
        }
    }

    /**
     * 解析段落
     */
    private void parseParagraphs(List<XWPFParagraph> paragraphs) {
        preHandleParagraphs(paragraphs);
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
                    if (beginIndex == 1) {
                        beginIndex = runText.indexOf(Commons.DOLLAR) + 1;
                    }
                    int endIndex = runText.indexOf(Commons.RIGHT_BRACKETS);
                    if (beginIndex != 1 && endIndex != -1 && beginIndex < endIndex) {
                        // 此 run 包含表达式
                        String exp = runText.substring(beginIndex, endIndex);
                        tmpRun.setText(runText.replace(Commons.LEFT_BRACKETS + exp + Commons.RIGHT_BRACKETS, getExpressionValue(exp)), 0);
                    } else if (beginIndex != 1) {
                        // 存在 ${ 或者 $
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
     * 预处理 paragraphs
     *
     * @param paragraphs 待预处理的 paragraphs 的 List
     */
    private void preHandleParagraphs(List<XWPFParagraph> paragraphs) {
        int paragraphSize = paragraphs.size();
        for (int pIndex = 0; pIndex < paragraphSize; pIndex++) {
            XWPFParagraph paragraph = paragraphs.get(pIndex);
            // 段落的 text
            String paragraphText = paragraph.getText();
            // 存放分隔开的内容的 List
            List<String> splitTextList = new ArrayList<>();
            splitConfigurationArray(paragraphText, splitTextList);

            // run 中的多余内容, 需要加到下一个 run 里面的
            StringBuilder moreContent = new StringBuilder();
            // 这个 p 指向 splitTextList 中的元素的索引
            AtomicInteger p = new AtomicInteger(0);
            List<XWPFRun> runs = paragraph.getRuns();
            // runs 的 size, 可及时改变
            int runSize = runs.size();
            for (int runIndex = 0; runIndex < runSize; runIndex++) {
                XWPFRun run = runs.get(runIndex);
                // 文本位置
                int runTextPosition = run.getTextPosition() == -1 ? 0 : run.getTextPosition();
                String runText = run.getText(runTextPosition);
                // 第 p 个元素
                String pText = splitTextList.get(p.get());
                if (pText.contains(runText)) {
                    // 如果 pText 包括 runText, 说明 runText 很有可能是一个纯文本
                    // 因为 splitTextList 是已经分好了的 List 了, 不可能出现配置和纯文本混淆的情况
                    // 此时需要将这个纯文本 runText 从 pText 中去除, 并将 pText 重新放入 splitTextList 中
                    pText = pText.substring(pText.indexOf(runText) + runText.length());
                    splitTextList.set(p.get(), pText);
                } else {
                    // 到这里的 else 分支后, runText 并不是纯文本了, runText 可能是一下几种格式
                    // nameis${user, nameis${username}, ${user, ${username}123 等四种格式
                    // 首先将可能存在的配置取出来
                    Matcher matcher = Document.SingletonPattern.pattern.matcher(runText);
                    if (matcher.find()) {
                        // 有完整的配置
                        String configuration = matcher.group();
                        // 配置的位置
                        int configIndex = runText.indexOf(configuration);
                        if (runText.startsWith(configuration)) {
                            // 如果 runText 开始就是配置
                            run.setText(runText.substring(0, configIndex + configuration.length()), runTextPosition);
                            p.incrementAndGet();
                            // 然后将剩下的部分新建一个 run 并放到里面
                            XWPFRun newRun = paragraph.insertNewRun(runIndex + 1);
                            copyStyle(run, newRun);
                            newRun.setText(runText.substring(configIndex + configuration.length()), newRun.getTextPosition());
                            runSize++;
                        } else {
                            // 如果开始不是配置
                            String tmpRunText = runText.substring(0, configIndex);
                            run.setText(tmpRunText, runTextPosition);
                            p.incrementAndGet();
                            // 然后将剩下的部分新建一个 run 并放到里面
                            XWPFRun newRun = paragraph.insertNewRun(runIndex + 1);
                            copyStyle(run, newRun);
                            newRun.setText(runText.substring(configIndex), newRun.getTextPosition());
                            runSize++;
                        }
                    } else {
                        // 走到这里说明 runText 中并没有完整的配置, 此时需要将 runText 中对应的 pText 的值删除掉并重新 set 进 run 中
                        // pTextBeginIndex 表示从哪个位置可以截取 runText
                        int pTextBeginIndex = runText.indexOf(pText) + pText.length();
                        String tmpRunText = runText.substring(0, pTextBeginIndex);
                        run.setText(tmpRunText, runTextPosition);
                        // 需要遍历 splitTextList 的下一个元素了
                        p.incrementAndGet();
                        // 然后将剩下的部分新建一个 run 并放到里面
                        XWPFRun newRun = paragraph.insertNewRun(runIndex + 1);
                        copyStyle(run, newRun);
                        newRun.setText(runText.substring(pTextBeginIndex), newRun.getTextPosition());
                        runSize++;
                    }
                }
            }
        }
    }

    /**
     * 将 run 的样式拷贝到 newRun 上面
     *
     * @param run    run
     * @param newRun newRun
     */
    private void copyStyle(XWPFRun run, XWPFRun newRun) {
        // 字体颜色
        newRun.setColor(run.getColor());
        // 字体
        newRun.setFontFamily(run.getFontFamily());
        // 字体大小
        newRun.setFontSize(run.getFontSize());
    }

    /**
     * 根据配置拆分为 String 数组
     *
     * @param text text
     * @return 拆分后的数组
     */
    private void splitConfigurationArray(String text, List<String> list) {
        Matcher matcher = Document.SingletonPattern.pattern.matcher(text);
        if (!matcher.find()) {
            list.add(text);
            return;
        }
        String validText = matcher.group();
        // 这段内容在总的内容中的索引
        int index = text.indexOf(validText);
        list.add(text.substring(0, index));
        list.add(validText);
        splitConfigurationArray(text.substring(index + validText.length()), list);
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
    public static List<XWPFParagraph> getParsingParagraphs(List<XWPFParagraph> paragraphs) {
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
     * 搜索可以解析的表格 {@link XWPFTable}
     *
     * @param tables 全部的表格对象的 List
     * @return 待解析的 List
     */
    public static List<XWPFTable> getParsingTables(List<XWPFTable> tables) {
        List<XWPFTable> waitParseList = new ArrayList<>(tables.size());
        for (XWPFTable table : tables) {
            // 表格的文本
            String text = table.getText();
            if (!Document.existsExpress(text, true)) {
                // 宽松的验证符合正则表达式的文本
                continue;
            }
            waitParseList.add(table);
        }
        return waitParseList;
    }

    /**
     * 单例模式获取 {@link JexlEngine} 对象
     */
    private static class SingleJexlEngine {
        static JexlEngine jexlEngine = new JexlBuilder().create();
    }
}
