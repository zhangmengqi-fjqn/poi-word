package com.melon.word;

import org.apache.commons.lang3.StringUtils;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author zhaokai
 * @date 2019-11-11
 */
public class Style {

    /**
     * @see CTStyle
     */
    private List<CTStyle> styleList;

    /**
     * style 的 name 对应的 style
     */
    private Map<String, CTStyle> nameMap;

    /**
     * 默认的段落样式的 name
     */
    public static final String NORMAL = "Normal";

    public Style(List<CTStyle> styles) {
        this.styleList = styles;
    }

    public CTStyle getByName(String name) {
        if (StringUtils.isEmpty(name)) {
            // 为空直接返回
            return null;
        }
        if (nameMap == null) {
            // 第一次的话需要初始化一下
            nameMap = new HashMap<>(16);
            for (CTStyle ctStyle : styleList) {
                // 将 name 的值和 style 本身放进去
                CTString styleName = ctStyle.getName();
                nameMap.put(styleName.getVal(), ctStyle);
            }
        }
        return nameMap.get(name);
    }

    public CTStyle getDefaultNormalStyle() {
        CTStyle ctStyle = getByName(NORMAL);
        // 为空则直接返回
        if (ctStyle == null) {
            return null;
        }
        // 设置了默认并且默认值为 1 才能返回, 否则返回 null
        return ctStyle.isSetDefault() && Objects.equals("1", ctStyle.getDefault().toString()) ? ctStyle : null;
    }
}
