package com.motcs.utils.poi.enums;

import lombok.Getter;

/**
 * @author <a href="https://github.com/motcs">motcs</a>
 * @since 2024-03-28 星期四
 */
@Getter
public enum Color {

    /**
     * 黑色
     */
    BLACK("000000"),

    /**
     * 白色
     */
    WHITE("FFFFFF"),

    /**
     * 红色
     */
    RED("FF0000"),

    /**
     * 绿色
     */
    GREEN("00FF00"),

    /**
     * 蓝色色
     */
    BLUE("0000FF");


    private final String value;

    Color(String value) {
        this.value = value;
    }

}
