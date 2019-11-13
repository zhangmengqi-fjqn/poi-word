package com.melon.word.common;

import java.util.UUID;

/**
 * @author zhaokai
 * @date 2019-11-13
 */
public class UUIDTools {

    private UUIDTools() {
    }

    /**
     * 获取 UUID，不要 '-'，全部转换为小写
     *
     * @return result
     */
    public static String getLowerUUID() {
        return UUID.randomUUID().toString().replace("-", "").toLowerCase();
    }
}
