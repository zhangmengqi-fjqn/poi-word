package com.melon.word;

import org.apache.commons.lang3.StringUtils;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

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
     * 默认的 ppr default
     */
    private CTPPrDefault ctpPrDefault;

    /**
     * 默认的 rpr default
     */
    private CTRPrDefault ctrPrDefault;

    /**
     * style 的 name 对应的 style
     */
    private Map<String, CTStyle> nameMap;

    /**
     * style 的 id 对应的 style
     */
    private Map<String, CTStyle> idMap;

    /**
     * 默认的段落样式的 name
     */
    public static final String NORMAL = "Normal";

    /**
     * 默认的 table 样式的 name
     */
    public static final String NORMAL_TABLE = "Normal Table";

    public Style(CTStyles ctStyles) {
        this.styleList = ctStyles.getStyleList();
        // 初始化 nameMap
        nameMap = new HashMap<>(16);
        idMap = new HashMap<>(16);
        for (CTStyle ctStyle : styleList) {
            // 将 name 的值和 style 本身放进去
            nameMap.put(ctStyle.getName().getVal(), ctStyle);
            // 将 styleId 和 style 本身存入
            idMap.put(ctStyle.getStyleId(), ctStyle);
        }
        CTDocDefaults docDefaults = ctStyles.getDocDefaults();
        if (docDefaults != null) {
            // 不为空设置默认的 ppr 和 rpr
            this.ctpPrDefault = docDefaults.getPPrDefault();
            this.ctrPrDefault = docDefaults.getRPrDefault();
        }
    }

    /**
     * 根据 styleId 获取 style
     *
     * @param styleId styleId
     * @return 对应的 style，没有返回 null
     */
    public CTStyle getByStyleId(String styleId) {
        if (StringUtils.isBlank(styleId)) {
            return null;
        }
        return idMap.get(styleId);
    }

    public CTStyle getByName(String name) {
        if (StringUtils.isEmpty(name)) {
            // 为空直接返回
            return null;
        }
        return nameMap.get(name);
    }

    /**
     * 获取默认的 style
     *
     * @return 默认的 style
     */
    public CTStyle getDefaultNormalStyle() {
        CTStyle ctStyle = getByName(NORMAL);
        // 为空则直接返回
        if (ctStyle == null) {
            return null;
        }
        // 设置了默认并且默认值为 1 才能返回, 否则返回 null
        return ctStyle.isSetDefault() && Objects.equals("1", ctStyle.getDefault().toString()) ? ctStyle : null;
    }

    /**
     * 获取默认的 table style
     *
     * @return 默认的 table style
     */
    public CTStyle getDefaultTableStyle() {
        CTStyle ctStyle = getByName(NORMAL_TABLE);
        // 为空则直接返回
        if (ctStyle == null) {
            return null;
        }
        // 设置了默认并且默认值为 1 才能返回, 否则返回 null
        return ctStyle.isSetDefault() && Objects.equals("1", ctStyle.getDefault().toString()) ? ctStyle : null;
    }

    /**
     * 获取默认的段落的 style
     *
     * @return 段落的默认的 style {@link CTPPr}
     */
    public CTPPr getDefaultCTPPr() {
        CTStyle defaultNormalStyle = getDefaultNormalStyle();
        if (defaultNormalStyle != null) {
            return defaultNormalStyle.getPPr();
        }
        // 这个只能获取默认的 ppr default 了
        if (ctpPrDefault == null) {
            return null;
        }
        return ctpPrDefault.getPPr();
    }

    /**
     * 获取默认的 run 的 style
     *
     * @return {@link org.apache.poi.xwpf.usermodel.XWPFRun}的默认的 style {@link CTRPr}
     */
    public CTRPr getDefaultCTRPr() {
        CTStyle defaultNormalStyle = getDefaultNormalStyle();
        if (defaultNormalStyle != null) {
            return defaultNormalStyle.getRPr();
        }
        // 这个只能获取默认的 ppr default 了
        if (ctrPrDefault == null) {
            return null;
        }
        return ctrPrDefault.getRPr();
    }

}
