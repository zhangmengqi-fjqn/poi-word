package com.melon.word.common;

/**
 * 通用常量
 *
 * @author zhaokai
 * @date 2019-10-15
 */
public interface CommonConstants {


    /**
     * 使用 ${} 包裹着的的内容的匹配正则表达式
     */
    String DOLLAR_REGEX = "\\$\\{.*?\\}";


    /**
     * 配置的左边部分的括号
     */
    String LEFT_BRACKETS = "${";


    /**
     * 配置的右边部分的括号
     */
    String RIGHT_BRACKETS = "}";


    /**
     * 字符串类型的换行符
     */
    String NEW_LINE_STRING = "\n";

    /**
     * 美元符号
     */
    String DOLLAR = "$";

    /**
     * 人民币符号
     */
    String RMB = "￥";

}
