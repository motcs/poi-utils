package com.motcs.utils.poi.enums;

import lombok.Getter;

/**
 * @author <a href="https://github.com/motcs">motcs</a>
 * @since 2024-03-28 星期四
 */@Getter
public enum Footer {


    /**
     * 黑色
     */
    BEGIN("begin"),

    /**
     * 红色
     */
    END("end"),

    PAGE("PAGE  \\* MERGEFORMAT"),

    /**
     * 用于指示文本之间的空间应该保持不变，而不进行额外的调整或压缩<br>
     * 这可能适用于文本布局或格式化中，以确保特定文本元素之间的间距保持不变
     */
    PRESERVE("preserve");


    private final String value;

    Footer(String value) {
        this.value = value;
    }

}
