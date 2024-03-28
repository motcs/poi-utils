package com.motcs.utils.poi.enums;

import lombok.Getter;

/**
 * @author <a href="https://github.com/motcs">motcs</a>
 * @since 2024-03-28 星期四
 */
@Getter
public enum ChineseFont {

    /**
     * 宋体
     */
    SONG_TI("宋体"),

    /**
     * 仿宋GB2312
     */
    FANG_SONG_GB2312("仿宋GB2312"),

    /**
     * 黑体
     */
    HEI_TI("黑体"),

    /**
     * 楷体
     */
    KAI_TI("楷体"),

    /**
     * 隶书
     */
    LI_SHU("隶书"),

    /**
     * 幼圆
     */
    YOU_YUAN("幼圆");


    private final String fontName;

    ChineseFont(String fontName) {
        this.fontName = fontName;
    }

}

